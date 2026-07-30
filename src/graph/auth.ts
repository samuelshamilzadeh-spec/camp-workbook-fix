import {
  ClientSecretCredential,
  ManagedIdentityCredential,
  type TokenCredential,
} from '@azure/identity';
import type { RuntimeConfig } from '../config';

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

export interface TokenProvider {
  getToken(): Promise<string>;
}

/**
 * App-only auth. In Azure prefer a managed identity so there is no secret to
 * rotate; locally a client secret is the only option.
 *
 * Either way the app registration carries `Sites.Selected` and nothing broader,
 * and the site-level grant is what actually scopes access to the one site
 * holding the workbook. `Sites.Selected` with no site grant can read nothing,
 * which is the desired default.
 */
export function createCredential(config: RuntimeConfig): TokenCredential {
  if (config.useManagedIdentity) {
    return config.managedIdentityClientId
      ? new ManagedIdentityCredential({ clientId: config.managedIdentityClientId })
      : new ManagedIdentityCredential();
  }
  if (!config.clientSecret) {
    throw new Error(
      'AZURE_CLIENT_SECRET is required unless AZURE_USE_MANAGED_IDENTITY=true.',
    );
  }
  return new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
}

/**
 * Caches the token in the worker process. At one invocation every 5 seconds a
 * fresh token request per cycle would be roughly 17,000 needless calls a day.
 */
export function createTokenProvider(credential: TokenCredential): TokenProvider {
  let cached: { token: string; expiresOnMs: number } | undefined;
  let inFlight: Promise<string> | undefined;

  const SKEW_MS = 5 * 60 * 1000;

  return {
    async getToken(): Promise<string> {
      if (cached && cached.expiresOnMs - SKEW_MS > Date.now()) {
        return cached.token;
      }
      if (inFlight) return inFlight;

      inFlight = (async () => {
        try {
          const result = await credential.getToken(GRAPH_SCOPE);
          if (!result) throw new Error('Credential returned no token for Microsoft Graph.');
          cached = { token: result.token, expiresOnMs: result.expiresOnTimestamp };
          return result.token;
        } finally {
          inFlight = undefined;
        }
      })();

      return inFlight;
    },
  };
}
