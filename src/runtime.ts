/**
 * Deno runtime helpers.
 */

const BASE64_CHUNK_SIZE = 0x8000;

export function encodeBase64(bytes: Uint8Array): string {
  let binary = "";

  for (let offset = 0; offset < bytes.length; offset += BASE64_CHUNK_SIZE) {
    const chunk = bytes.subarray(offset, offset + BASE64_CHUNK_SIZE);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary);
}

export function isRemotePath(path: string): boolean {
  return path.startsWith("http://") || path.startsWith("https://");
}

export async function readPathBytes(path: string): Promise<Uint8Array> {
  if (!isRemotePath(path)) return await Deno.readFile(path);

  const response = await fetch(path);

  if (!response.ok) {
    throw new Error(
      `ERROR! Unable to load media (${response.status} ${response.statusText}): ${path}`,
    );
  }

  return new Uint8Array(await response.arrayBuffer());
}

export async function readPathAsBase64(path: string): Promise<string> {
  return encodeBase64(await readPathBytes(path));
}
