import { OAuth2Client } from 'google-auth-library';
import { google, drive_v3 } from 'googleapis';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import { initializeOAuth2Client } from './client.js';
import { TokenManager } from './tokenManager.js';

export interface AccountConfig {
  alias: string;
}

export interface AccountsConfig {
  accounts: Record<string, AccountConfig>;
  defaultAccount: string;
}

export interface AccountServices {
  authClient: OAuth2Client;
  drive: drive_v3.Drive;
  email: string;
  alias?: string;
}

export class AccountManager {
  private config: AccountsConfig | null = null;
  private configLoaded = false;
  private serviceCache: Map<string, AccountServices> = new Map();
  private multiAccountMode = false;

  private getConfigPath(): string {
    const configHome = process.env.XDG_CONFIG_HOME ||
      path.join(os.homedir(), '.config');
    return path.join(configHome, 'google-drive-mcp', 'accounts.json');
  }

  private getAccountTokenPath(email: string): string {
    const configHome = process.env.XDG_CONFIG_HOME ||
      path.join(os.homedir(), '.config');
    return path.join(configHome, 'google-drive-mcp', 'accounts', email, 'token.json');
  }

  private async loadConfig(): Promise<void> {
    if (this.configLoaded) return;
    this.configLoaded = true;

    try {
      const configPath = this.getConfigPath();
      const content = await fs.readFile(configPath, 'utf-8');
      this.config = JSON.parse(content) as AccountsConfig;
      this.multiAccountMode = true;
      console.error(`Multi-account mode: loaded ${Object.keys(this.config.accounts).length} accounts`);
    } catch {
      // No config file = single-account mode
      this.config = null;
      this.multiAccountMode = false;
    }
  }

  /**
   * Resolve an account parameter (alias, email, or undefined) to a canonical email.
   * In single-account mode, always returns undefined (use legacy auth).
   */
  resolveAccount(param?: string): string | undefined {
    if (!this.multiAccountMode || !this.config) return undefined;

    if (!param) {
      return this.config.defaultAccount;
    }

    // Check if param is already a known email
    if (this.config.accounts[param]) {
      return param;
    }

    // Check if param is an alias
    for (const [email, cfg] of Object.entries(this.config.accounts)) {
      if (cfg.alias === param) {
        return email;
      }
    }

    // Treat as email even if not in config (allows ad-hoc usage)
    return param;
  }

  /**
   * Get authenticated services for a given account.
   * In single-account mode (no accounts.json), returns null to signal
   * the caller should use legacy globals.
   */
  async getServices(account?: string): Promise<AccountServices | null> {
    await this.loadConfig();

    if (!this.multiAccountMode) {
      return null; // Caller should use legacy globals
    }

    const email = this.resolveAccount(account);
    if (!email) {
      return null; // Shouldn't happen when multiAccountMode is true, but be safe
    }

    // Check cache
    const cached = this.serviceCache.get(email);
    if (cached) {
      return cached;
    }

    // Init fresh services for this account
    const tokenPath = this.getAccountTokenPath(email);
    const oauth2Client = await initializeOAuth2Client();
    const tokenManager = new TokenManager(oauth2Client, tokenPath);

    if (!(await tokenManager.validateTokens())) {
      throw new Error(
        `Account "${email}" is not authenticated. ` +
        `Run: npx google-drive-mcp auth --account ${email}`
      );
    }

    const driveService = google.drive({ version: 'v3', auth: oauth2Client });

    const alias = this.config?.accounts[email]?.alias;
    const services: AccountServices = {
      authClient: oauth2Client,
      drive: driveService,
      email,
      alias,
    };

    this.serviceCache.set(email, services);
    return services;
  }

  /**
   * Returns a suffix to append to tool responses in multi-account mode.
   * Returns empty string in single-account mode.
   */
  formatAccountSuffix(services: AccountServices | null): string {
    if (!this.multiAccountMode || !services) return '';
    const aliasStr = services.alias ? ` (${services.alias})` : '';
    return `\n\n[Account: ${services.email}${aliasStr}]`;
  }

  isMultiAccount(): boolean {
    return this.multiAccountMode;
  }

  /**
   * Authenticate a specific account (used by CLI --account flag).
   */
  async authenticateAccount(email: string): Promise<void> {
    await this.loadConfig();

    const tokenPath = this.getAccountTokenPath(email);

    // Ensure directory exists
    const dir = path.dirname(tokenPath);
    await fs.mkdir(dir, { recursive: true });

    const oauth2Client = await initializeOAuth2Client();
    const { AuthServer } = await import('./server.js');
    const authServer = new AuthServer(oauth2Client, tokenPath);

    const success = await authServer.start(true);

    if (!success && !authServer.authCompletedSuccessfully) {
      throw new Error(
        `Authentication failed for ${email}. Check port availability (3000-3004) and try again.`
      );
    }

    if (authServer.authCompletedSuccessfully) {
      console.error(`\nAuthentication successful for ${email}!`);
      console.error(`Token saved to: ${tokenPath}`);
      return;
    }

    // Wait for browser callback
    console.error(`Waiting for authentication for ${email}...`);
    await new Promise<void>((resolve, reject) => {
      const intervalId = setInterval(async () => {
        if (authServer.authCompletedSuccessfully) {
          clearInterval(intervalId);
          await authServer.stop();
          console.error(`\nAuthentication completed for ${email}!`);
          console.error(`Token saved to: ${tokenPath}`);
          resolve();
        }
      }, 1000);

      // Timeout after 5 minutes
      setTimeout(() => {
        clearInterval(intervalId);
        reject(new Error(`Authentication timed out for ${email}`));
      }, 5 * 60 * 1000);
    });
  }
}
