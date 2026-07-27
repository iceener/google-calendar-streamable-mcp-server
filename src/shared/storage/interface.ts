export type ProviderTokens = {
  access_token: string;
  refresh_token?: string;
  expires_at?: number;
  scopes?: string[];
};

export type RsRecord = {
  rs_access_token: string;
  rs_refresh_token: string;
  provider: ProviderTokens;
  created_at: number;
};

export type Transaction = {
  codeChallenge: string;
  state?: string;
  scope?: string;
  createdAt: number;
  sid?: string;
  provider?: ProviderTokens;
};

export interface TokenStore {
  storeRsMapping(
    rsAccess: string,
    provider: ProviderTokens,
    rsRefresh?: string,
  ): Promise<RsRecord>;
  getByRsAccess(rsAccess: string): Promise<RsRecord | null>;
  getByRsRefresh(rsRefresh: string): Promise<RsRecord | null>;
  updateByRsRefresh(
    rsRefresh: string,
    provider: ProviderTokens,
    maybeNewRsAccess?: string,
  ): Promise<RsRecord | null>;
  saveTransaction(
    txnId: string,
    transaction: Transaction,
    ttlSeconds?: number,
  ): Promise<void>;
  getTransaction(txnId: string): Promise<Transaction | null>;
  deleteTransaction(txnId: string): Promise<void>;
  saveCode(code: string, txnId: string, ttlSeconds?: number): Promise<void>;
  getTxnIdByCode(code: string): Promise<string | null>;
  deleteCode(code: string): Promise<void>;
}
