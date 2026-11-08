const DATABASE_UNAVAILABLE_MESSAGE = 'Could not reach the database, please try again';

const databaseNetworkErrorNames = new Set([
  'MongoServerSelectionError',
  'MongoNetworkError',
  'MongoNetworkTimeoutError',
  'MongoOperationTimeoutError',
  'MongoNotConnectedError',
  'MongoTopologyClosedError',
]);

const networkErrorCodes = new Set([
  'ECONNREFUSED',
  'ECONNRESET',
  'EHOSTUNREACH',
  'ENETUNREACH',
  'ETIMEDOUT',
  'ERR_SSL_TLSV1_ALERT_INTERNAL_ERROR',
]);

export function getApiErrorMessage(error: unknown, fallback: string) {
  const pending = [error];
  const visited = new Set<unknown>();
  while (pending.length) {
    const current = pending.pop();
    if (!current || typeof current !== 'object' || visited.has(current)) continue;
    visited.add(current);

    if ('name' in current && typeof current.name === 'string' && databaseNetworkErrorNames.has(current.name)) {
      return DATABASE_UNAVAILABLE_MESSAGE;
    }
    if ('code' in current && typeof current.code === 'string' && networkErrorCodes.has(current.code)) {
      return DATABASE_UNAVAILABLE_MESSAGE;
    }
    if ('cause' in current) pending.push(current.cause);
  }
  return fallback;
}

export function requestExceedsBodyLimit(request: Request, limitBytes = 4 * 1024 * 1024) {
  const contentLength = Number(request.headers.get('content-length'));
  return Number.isFinite(contentLength) && contentLength > limitBytes;
}
