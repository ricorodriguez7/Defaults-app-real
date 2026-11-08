export const MAX_IMPORT_REQUEST_BYTES = 4 * 1024 * 1024;
export const MAX_IMPORT_SIZE_MESSAGE = 'Import requests must be under 4 MB. Split the spreadsheet and upload smaller files.';

export class RequestBodyError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
    this.name = 'RequestBodyError';
  }
}

export async function readJsonBody<T>(request: Request): Promise<T> {
  const contentLength = Number(request.headers.get('content-length') ?? 0);
  if (contentLength > MAX_IMPORT_REQUEST_BYTES) {
    throw new RequestBodyError(MAX_IMPORT_SIZE_MESSAGE, 413);
  }

  const body = await request.text();
  if (Buffer.byteLength(body, 'utf8') > MAX_IMPORT_REQUEST_BYTES) {
    throw new RequestBodyError(MAX_IMPORT_SIZE_MESSAGE, 413);
  }

  try {
    return JSON.parse(body) as T;
  } catch {
    throw new RequestBodyError('The upload is not valid JSON. Please try the spreadsheet again.', 400);
  }
}
