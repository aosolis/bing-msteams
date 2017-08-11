import * as edge from "edge";
import * as path from "path";

export interface GetOptions {
  storeName?: string;
  storeLocation?: string;
}

export interface X509Certificate {
  pem: string;
  subject: string;
  thumbprint: string;
  issuer: string;
}

export type CertificatesCallback = (error: Error|null, certificates?: X509Certificate[]) => void;

const getCerts = edge.func(path.join(__dirname, "get_certs.csx"));

function internal_get(options: GetOptions, callback: CertificatesCallback|true): void|X509Certificate[] {
  let params = {
    storeName: options.storeName || "",
    storeLocation: options.storeLocation || "",
    hasStoreName: !!options.storeName,
    hasStoreLocation: !!options.storeLocation,
  };
  return getCerts(params, callback);
}

export function get(options: GetOptions, callback?: CertificatesCallback): void|X509Certificate[] {
  if (callback) {
    return internal_get(options, callback);
  } else {
    return internal_get(options, true);
  }
}
