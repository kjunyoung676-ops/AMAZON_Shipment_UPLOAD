/**
 * FBA Inbound SP-API 라우트
 * POST /api/fba-inbound   { action, ...params }
 *
 * action 목록:
 *   createPlan            출하 플랜 생성
 *   pollOperation         비동기 작업 상태 조회
 *   generatePackingOptions  패킹 옵션 생성
 *   listPackingOptions    패킹 옵션 조회
 *   confirmPackingOption  패킹 옵션 확정
 *   setPackingInfo        박스 내용물 입력
 *   generatePlacementOptions  배치 옵션 생성
 *   listPlacementOptions  배치 옵션 조회
 *   confirmPlacement      FC 배치 확정 (수동)
 *   generateTransportOptions  운송 옵션 생성
 *   listTransportOptions  운송 옵션 조회
 *   confirmTransport      운송 옵션 확정 (비파트너 LTL)
 *   updateTracking        PRO 번호 입력
 *   getLabels             카톤 라벨 PDF URL
 *   getPalletLabels       파렛트 라벨 PDF URL
 *   getBol                BOL PDF URL
 */

import { NextRequest, NextResponse } from 'next/server'
import { spApiRequest, getOperationStatus } from '@/lib/sp-api'

const MP_ID = process.env.SP_API_MARKETPLACE_ID || 'A1VC38T7YXB528'  // JP

// ── 출고처 창고 목록 ─────────────────────────────────────────────────
export const CJ_WAREHOUSES = {
  saitama: {
    label:               '사이타마 창고',
    companyName:         process.env.SP_API_SOURCE_FULLNAME_SAITAMA  || '',
    addressLine1:        process.env.SP_API_SOURCE_ADDRESS1_SAITAMA  || '',
    addressLine2:        process.env.SP_API_SOURCE_ADDRESS2_SAITAMA  || '',
    city:                process.env.SP_API_SOURCE_CITY_SAITAMA      || '',
    countryCode:         'JP',
    postalCode:          process.env.SP_API_SOURCE_POSTAL_SAITAMA    || '',
    stateOrProvinceCode: process.env.SP_API_SOURCE_STATE_SAITAMA     || '',
    phoneNumber:         process.env.SP_API_SOURCE_PHONE_SAITAMA     || '',
  },
  osaka: {
    label:               '오사카 창고',
    companyName:         process.env.SP_API_SOURCE_FULLNAME_OSAKA    || '',
    addressLine1:        process.env.SP_API_SOURCE_ADDRESS1_OSAKA    || '',
    addressLine2:        process.env.SP_API_SOURCE_ADDRESS2_OSAKA    || '',
    city:                process.env.SP_API_SOURCE_CITY_OSAKA        || '',
    countryCode:         'JP',
    postalCode:          process.env.SP_API_SOURCE_POSTAL_OSAKA      || '',
    stateOrProvinceCode: process.env.SP_API_SOURCE_STATE_OSAKA       || '',
    phoneNumber:         process.env.SP_API_SOURCE_PHONE_OSAKA       || '',
  },
} as const
export type WarehouseKey = keyof typeof CJ_WAREHOUSES

// ── 박스 그루핑 계산 ─────────────────────────────────────────────────
// 각 SKU의 총 수량을 cpp(carton per pallet)로 나눠 박스 단위 생성
interface SkuBoxInfo {
  msku: string          // Amazon Merchant SKU (realSku)
  quantity: number      // 총 수량
  cpp: number           // 카톤당 수량 (1박스 = cpp개)
  bx: number; by: number; bz: number  // 박스 치수 (cm)
  kg: number            // 박스 1개 무게 (kg)
}

function buildBoxes(
  items: SkuBoxInfo[],
  ownerMap?: Map<string, { labelOwner: string; prepOwner: string }>,
) {
  // 엑셀 업로드 포맷 기준:
  //   Units per box = 1  (박스 1개 = 판매단위 1개)
  //   Number of boxes    = 총 수량
  //   cpp = 팔레트당 박스 수 (박스 내용물과 무관)
  //
  // labelOwner / prepOwner 는 listPackingGroupItems 에서 읽어 온 값 사용
  // (createInboundPlan 에서 결정된 값과 반드시 일치해야 함)
  return items
    .filter(item => item.msku && item.quantity > 0)
    .map(item => {
      const owners = ownerMap?.get(item.msku)
      return {
        contentInformationSource: 'BOX_CONTENT_PROVIDED',
        dimensions: {
          height: item.bz,
          length: item.bx,
          width:  item.by,
          unitOfMeasurement: 'CM',
        },
        weight: { unit: 'KG', value: Math.round(item.kg * 100) / 100 },
        quantity: item.quantity,   // 박스 수 = 총 수량 (박스당 1단위)
        items: [{
          msku:       item.msku,
          quantity:   1,           // 박스당 1판매단위
          prepOwner:  owners?.prepOwner  ?? 'NONE',
          labelOwner: owners?.labelOwner ?? 'NONE',   // API 응답값 우선, 없으면 NONE
        }],
      }
    })
}

// ── 핸들러 ───────────────────────────────────────────────────────────
export async function POST(req: NextRequest) {
  try {
    const body = await req.json()
    const { action } = body as { action: string }

    // ── 1. 출하 플랜 생성 ─────────────────────────────────────────────
    // SKU마다 labelOwner 허용값이 다름 → SELLER로 먼저 시도,
    // 에러 파싱 후 NONE 필요한 SKU만 교정해서 재시도
    if (action === 'createPlan') {
      const { planName, items, warehouseKey } = body as {
        planName: string
        items: Array<{ msku: string; quantity: number }>
        warehouseKey: WarehouseKey
      }
      const warehouse = CJ_WAREHOUSES[warehouseKey] ?? CJ_WAREHOUSES.saitama
      const sourceAddr = {
        name:                warehouse.companyName,
        addressLine1:        warehouse.addressLine1,
        ...(warehouse.addressLine2 ? { addressLine2: warehouse.addressLine2 } : {}),
        city:                warehouse.city,
        countryCode:         warehouse.countryCode,
        postalCode:          warehouse.postalCode,
        stateOrProvinceCode: warehouse.stateOrProvinceCode,
        ...(warehouse.phoneNumber ? { phoneNumber: warehouse.phoneNumber } : {}),
      }

      // labelOwner 재정의 맵 (msku → 'NONE' | 'SELLER')
      const labelOwnerMap: Record<string, 'NONE' | 'SELLER'> = {}

      const buildItems = () => items.map(i => ({
        msku:       i.msku,
        quantity:   i.quantity,
        labelOwner: labelOwnerMap[i.msku] ?? 'SELLER',
        prepOwner:  'NONE' as const,
      }))

      // 1차 시도
      let result: unknown
      try {
        result = await spApiRequest('POST', '/inbound/fba/2024-03-20/inboundPlans', {}, {
          name: planName, destinationMarketplaces: [MP_ID],
          sourceAddress: sourceAddr, items: buildItems(),
        })
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)

        // "does not require labelOwner" → 해당 SKU는 NONE으로 교정
        const needsNone = [...msg.matchAll(/ERROR: ([^\s]+) does not require labelOwner/g)]
          .map(m => m[1])
        // "requires labelOwner" & NONE assigned → SELLER로 교정 (이미 기본값이지만 명시)
        const needsSeller = [...msg.matchAll(/ERROR: ([^\s]+) requires labelOwner but NONE/g)]
          .map(m => m[1])

        if (needsNone.length === 0 && needsSeller.length === 0) throw err

        needsNone.forEach(msku => { labelOwnerMap[msku] = 'NONE' })
        needsSeller.forEach(msku => { labelOwnerMap[msku] = 'SELLER' })

        // 2차 시도 (교정된 값으로)
        result = await spApiRequest('POST', '/inbound/fba/2024-03-20/inboundPlans', {}, {
          name: planName, destinationMarketplaces: [MP_ID],
          sourceAddress: sourceAddr, items: buildItems(),
        })
      }

      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 2. 비동기 작업 상태 폴링 ─────────────────────────────────────
    if (action === 'pollOperation') {
      const { operationId } = body as { operationId: string }
      const result = await getOperationStatus(operationId)
      return NextResponse.json({ ok: true, ...result })
    }

    // ── 3. 패킹 옵션 생성 ─────────────────────────────────────────────
    if (action === 'generatePackingOptions') {
      const { planId } = body as { planId: string }
      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/packingOptions`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 4. 패킹 옵션 조회 ─────────────────────────────────────────────
    if (action === 'listPackingOptions') {
      const { planId } = body as { planId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/packingOptions`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 5. 패킹 옵션 확정 ─────────────────────────────────────────────
    if (action === 'confirmPackingOption') {
      const { planId, packingOptionId } = body as { planId: string; packingOptionId: string }
      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/packingOptions/${packingOptionId}/confirmation`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 6. 박스 내용물 입력 ───────────────────────────────────────────
    // Amazon이 packing group별로 MSKU를 이미 배분해놓음
    // → 각 그룹의 items를 먼저 조회하여 해당 MSKU만 필터링 후 전송
    // 모든 그룹을 단일 요청의 packageGroupings 배열에 담아야 함
    if (action === 'setPackingInfo') {
      const { planId, packingGroupIds, items } = body as {
        planId: string
        packingGroupIds: string[]
        items: SkuBoxInfo[]
      }

      // 각 packing group의 items 조회 → MSKU 필터 + labelOwner/prepOwner 값 추출
      const packageGroupings = await Promise.all(
        packingGroupIds.map(async packingGroupId => {
          const pgRes = await spApiRequest(
            'GET',
            `/inbound/fba/2024-03-20/inboundPlans/${planId}/packingGroups/${packingGroupId}/items`,
          ) as { items?: Array<{ msku: string; labelOwner?: string; prepOwner?: string }> }

          // msku → {labelOwner, prepOwner} 맵 구성 (createInboundPlan 에서 결정된 값)
          const ownerMap = new Map<string, { labelOwner: string; prepOwner: string }>(
            (pgRes.items ?? []).map(i => [
              i.msku,
              { labelOwner: i.labelOwner ?? 'NONE', prepOwner: i.prepOwner ?? 'NONE' },
            ]),
          )

          const allowedMskus = new Set(ownerMap.keys())
          const filtered = allowedMskus.size > 0
            ? items.filter(i => allowedMskus.has(i.msku))
            : items   // fallback: 조회 실패 시 전체 전송

          return { packingGroupId, boxes: buildBoxes(filtered, ownerMap) }
        }),
      )

      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/packingInformation`,
        {},
        { packageGroupings },
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 7. 배치 옵션 생성 ─────────────────────────────────────────────
    if (action === 'generatePlacementOptions') {
      const { planId } = body as { planId: string }
      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/placementOptions`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 8. 배치 옵션 조회 ─────────────────────────────────────────────
    if (action === 'listPlacementOptions') {
      const { planId } = body as { planId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/placementOptions`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 9. FC 배치 확정 ───────────────────────────────────────────────
    if (action === 'confirmPlacement') {
      const { planId, placementOptionId } = body as {
        planId: string
        placementOptionId: string
      }
      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/placementOptions/${placementOptionId}/confirmation`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 10. 운송 옵션 생성 ────────────────────────────────────────────
    if (action === 'generateTransportOptions') {
      const { planId, shipmentId, readyToShipWindow } = body as {
        planId: string
        shipmentId: string
        readyToShipWindow: { start: string; end: string }
      }
      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/transportationOptions`,
        {},
        { readyToShipWindow },
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 11. 운송 옵션 조회 ────────────────────────────────────────────
    if (action === 'listTransportOptions') {
      const { planId, shipmentId } = body as { planId: string; shipmentId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/transportationOptions`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 12. 운송 옵션 확정 (비파트너 LTL/FTL) ────────────────────────
    if (action === 'confirmTransport') {
      const { planId, shipmentId, transportOptionId } = body as {
        planId: string
        shipmentId: string
        transportOptionId: string
      }
      const result = await spApiRequest(
        'POST',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/transportationOptions/${transportOptionId}/confirmation`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 13. 팔레트 정보 등록 (LTL/FTL) ─────────────────────────────────
    // palletCount: 클라이언트가 listShipmentItems + master.cpp로 자동 계산해서 전달
    // carrierName: 하드코딩 금지 → env var SP_API_CARRIER_NAME 서버에서 직접 읽음
    if (action === 'updateTracking') {
      const { planId, shipmentId, palletCount, proNumber } = body as {
        planId: string
        shipmentId: string
        palletCount?: number
        proNumber?: string   // 실제 PRO 번호는 출하 후 업데이트 — 미제공 시 생략
      }
      const carrierName = process.env.SP_API_CARRIER_NAME || ''
      const result = await spApiRequest(
        'PUT',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/trackingDetails`,
        {},
        {
          ltlTrackingDetail: {
            ...(proNumber ? { freightBillNumber: proNumber } : {}),
            ...(palletCount ? { palletCount } : {}),
          },
          ...(carrierName ? { carrierName } : {}),
        },
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 13-b. Shipment 아이템 조회 (팔레트 수 계산용) ─────────────────
    if (action === 'listShipmentItems') {
      const { planId, shipmentId } = body as { planId: string; shipmentId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/items`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 14. 카톤 라벨 URL ─────────────────────────────────────────────
    if (action === 'getLabels') {
      const { planId, shipmentId, pageType, labelType } = body as {
        planId: string
        shipmentId: string
        pageType?: string
        labelType?: string
      }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/labels`,
        {
          pageType:  pageType  || 'PLAIN_PAPER',
          labelType: labelType || 'UNIQUE',
        },
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 15. 파렛트 라벨 URL ───────────────────────────────────────────
    if (action === 'getPalletLabels') {
      const { planId, shipmentId } = body as { planId: string; shipmentId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/labels`,
        {
          pageType:  'ZA_LABEL_PALLETS',
          labelType: 'UNIQUE',
        },
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 16. BOL ───────────────────────────────────────────────────────
    if (action === 'getBol') {
      const { planId, shipmentId } = body as { planId: string; shipmentId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments/${shipmentId}/billOfLading`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── 플랜 상세 조회 ────────────────────────────────────────────────
    if (action === 'getPlan') {
      const { planId } = body as { planId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    // ── Shipment 목록 조회 (confirmPackingOption 후 shipmentId 취득) ──
    if (action === 'listShipments') {
      const { planId } = body as { planId: string }
      const result = await spApiRequest(
        'GET',
        `/inbound/fba/2024-03-20/inboundPlans/${planId}/shipments`,
      )
      return NextResponse.json({ ok: true, ...(result as object) })
    }

    return NextResponse.json({ ok: false, error: `알 수 없는 action: ${action}` }, { status: 400 })

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    console.error('[fba-inbound]', msg)
    return NextResponse.json({ ok: false, error: msg }, { status: 500 })
  }
}
