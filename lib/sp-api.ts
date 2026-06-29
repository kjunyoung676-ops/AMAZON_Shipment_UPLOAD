/**
 * SP-API 인증 + 요청 헬퍼
 * LWA 토큰 취득 → STS AssumeRole → SigV4 서명
 */
import crypto from 'crypto'

// ── 상수 ────────────────────────────────────────────────────────────
const SP_ENDPOINT = 'https://sellingpartnerapi-fe.amazon.com'
const SP_REGION   = 'us-west-2'
const SP_SERVICE  = 'execute-api'
const STS_HOST    = 'sts.amazonaws.com'
const STS_REGION  = 'us-east-1'

// ── 토큰 캐시 ───────────────────────────────────────────────────────
let lwaCache: { token: string; expiresAt: number } | null = null
let stsCache: {
  accessKeyId: string
  secretAccessKey: string
  sessionToken: string
  expiresAt: number
} | null = null

// ── LWA 액세스 토큰 ─────────────────────────────────────────────────
export async function getLwaToken(): Promise<string> {
  if (lwaCache && Date.now() < lwaCache.expiresAt - 60_000) {
    return lwaCache.token
  }
  const res = await fetch('https://api.amazon.com/auth/o2/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type:    'refresh_token',
      refresh_token: process.env.SP_API_REFRESH_TOKEN!,
      client_id:     process.env.SP_API_CLIENT_ID!,
      client_secret: process.env.SP_API_CLIENT_SECRET!,
    }).toString(),
  })
  if (!res.ok) throw new Error(`LWA 토큰 오류: ${await res.text()}`)
  const d = await res.json()
  lwaCache = { token: d.access_token, expiresAt: Date.now() + d.expires_in * 1000 }
  return lwaCache.token
}

// ── SigV4 헬퍼 ──────────────────────────────────────────────────────
function hmac(key: Buffer | string, data: string): Buffer {
  return crypto.createHmac('sha256', key).update(data, 'utf8').digest()
}
function sha256Hex(data: string): string {
  return crypto.createHash('sha256').update(data, 'utf8').digest('hex')
}
function signingKey(secret: string, date: string, region: string, svc: string): Buffer {
  return hmac(hmac(hmac(hmac(`AWS4${secret}`, date), region), svc), 'aws4_request')
}

interface SignParams {
  method: string
  host: string
  path: string
  queryStr: string
  headers: Record<string, string>   // 소문자 키
  body: string
  accessKeyId: string
  secretKey: string
  sessionToken?: string
  region: string
  service: string
  amzDate: string
  dateStamp: string
}

function buildAuthHeader(p: SignParams): string {
  const sortedHdrKeys = Object.keys(p.headers).sort()
  const signedHeaders  = sortedHdrKeys.join(';')
  const canonicalHeaders =
    sortedHdrKeys.map(k => `${k}:${p.headers[k]}`).join('\n') + '\n'

  const payloadHash = sha256Hex(p.body)
  const canonicalRequest = [
    p.method, p.path, p.queryStr,
    canonicalHeaders, signedHeaders, payloadHash,
  ].join('\n')

  const scope  = `${p.dateStamp}/${p.region}/${p.service}/aws4_request`
  const strToSign = ['AWS4-HMAC-SHA256', p.amzDate, scope, sha256Hex(canonicalRequest)].join('\n')
  const sig = crypto
    .createHmac('sha256', signingKey(p.secretKey, p.dateStamp, p.region, p.service))
    .update(strToSign).digest('hex')

  return `AWS4-HMAC-SHA256 Credential=${p.accessKeyId}/${scope}, SignedHeaders=${signedHeaders}, Signature=${sig}`
}

// ── STS AssumeRole ───────────────────────────────────────────────────
// SP_API_ROLE_ARN 미설정 또는 user ARN인 경우 → IAM User 직접 사용
function isValidRoleArn(arn: string | undefined): boolean {
  return !!arn && arn.includes(':role/')
}

async function assumeRole(): Promise<typeof stsCache & object> {
  if (stsCache && Date.now() < stsCache.expiresAt - 60_000) return stsCache!

  const roleArn   = process.env.SP_API_ROLE_ARN
  const accessKey = process.env.SP_API_AWS_ACCESS_KEY!
  const secretKey = process.env.SP_API_AWS_SECRET_KEY!

  // Role ARN 없거나 잘못된 경우 → IAM User 자격증명 직접 반환 (STS 생략)
  if (!isValidRoleArn(roleArn)) {
    return {
      accessKeyId:     accessKey,
      secretAccessKey: secretKey,
      sessionToken:    '',
      expiresAt:       Date.now() + 3_600_000,
    }
  }

  const now       = new Date()
  const amzDate   = now.toISOString().replace(/[:-]|\.\d{3}/g, '').slice(0, 15) + 'Z'
  const dateStamp = amzDate.slice(0, 8)

  const body = new URLSearchParams({
    Action:          'AssumeRole',
    RoleArn:         roleArn,
    RoleSessionName: 'SpApiSession',
    DurationSeconds: '3600',
    Version:         '2011-06-15',
  }).toString()

  const headers: Record<string, string> = {
    'content-type': 'application/x-www-form-urlencoded',
    'host':         STS_HOST,
    'x-amz-date':  amzDate,
  }

  const auth = buildAuthHeader({
    method: 'POST', host: STS_HOST, path: '/', queryStr: '',
    headers, body,
    accessKeyId: accessKey, secretKey,
    region: STS_REGION, service: 'sts',
    amzDate, dateStamp,
  })

  const res = await fetch(`https://${STS_HOST}/`, {
    method: 'POST',
    headers: { ...headers, Authorization: auth },
    body,
  })
  const xml = await res.text()
  if (!res.ok) throw new Error(`STS AssumeRole 실패: ${xml}`)

  const get = (tag: string) => xml.match(new RegExp(`<${tag}>([^<]+)</${tag}>`))?.[1] ?? ''
  stsCache = {
    accessKeyId:     get('AccessKeyId'),
    secretAccessKey: get('SecretAccessKey'),
    sessionToken:    get('SessionToken'),
    expiresAt:       new Date(get('Expiration') || Date.now() + 3_600_000).getTime(),
  }
  return stsCache
}

// ── SP-API 요청 (메인 함수) ──────────────────────────────────────────
export async function spApiRequest(
  method: string,
  path: string,
  queryParams: Record<string, string> = {},
  body: unknown = null,
): Promise<unknown> {
  const [lwaToken, creds] = await Promise.all([getLwaToken(), assumeRole()])

  const now       = new Date()
  const amzDate   = now.toISOString().replace(/[:-]|\.\d{3}/g, '').slice(0, 15) + 'Z'
  const dateStamp = amzDate.slice(0, 8)
  const bodyStr   = body ? JSON.stringify(body) : ''

  const queryStr = Object.entries(queryParams)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&')

  const host = new URL(SP_ENDPOINT).hostname
  const headers: Record<string, string> = {
    'content-type':       'application/json',
    'host':               host,
    'x-amz-access-token': lwaToken,
    'x-amz-date':         amzDate,
    // sessionToken이 없으면 헤더 제외 (IAM User 직접 사용 시)
    ...(creds.sessionToken ? { 'x-amz-security-token': creds.sessionToken } : {}),
  }

  const auth = buildAuthHeader({
    method, host, path, queryStr,
    headers, body: bodyStr,
    accessKeyId: creds.accessKeyId,
    secretKey:   creds.secretAccessKey,
    sessionToken: creds.sessionToken,
    region: SP_REGION, service: SP_SERVICE,
    amzDate, dateStamp,
  })

  const url = `${SP_ENDPOINT}${path}${queryStr ? '?' + queryStr : ''}`
  const res = await fetch(url, {
    method,
    headers: { ...headers, Authorization: auth },
    body: bodyStr || undefined,
  })

  const text = await res.text()
  const json = text ? JSON.parse(text) : null
  if (!res.ok) throw new Error(`SP-API 오류 (${res.status}) ${method} ${path}: ${text}`)
  return json
}

// ── 유틸: 비동기 작업 상태 조회 ─────────────────────────────────────
export async function getOperationStatus(operationId: string): Promise<{
  operationStatus: 'IN_PROGRESS' | 'SUCCESS' | 'FAILED'
  operationProblems?: Array<{ code: string; message: string; severity: string }>
}> {
  const result = await spApiRequest(
    'GET',
    `/inbound/fba/2024-03-20/operations/${operationId}`,
  )
  return result as {
    operationStatus: 'IN_PROGRESS' | 'SUCCESS' | 'FAILED'
    operationProblems?: Array<{ code: string; message: string; severity: string }>
  }
}
