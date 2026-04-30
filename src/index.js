#!/usr/bin/env node

/**
 * Agent Index Filesystem MCP Server — OneDrive/SharePoint adapter
 *
 * Entry point. Loads config from AIFS_CONFIG_PATH, initializes the
 * OneDrive adapter, and starts the MCP server on stdio.
 */

import { initEnvironment, loadConfig, startServer } from '@agent-index/filesystem';
import { OneDriveAdapter } from './adapters/onedrive.js';

async function main() {
  // Detect proxy environment and configure TLS before any HTTP calls.
  initEnvironment();

  let config;
  try {
    config = await loadConfig();
  } catch (err) {
    console.error(`[aifs] Config error: ${err.message}`);
    process.exit(1);
  }

  if (config.backend !== 'onedrive') {
    console.error(
      `[aifs] This package (@agent-index/filesystem-onedrive) only supports the "onedrive" backend. ` +
      `Config specifies "${config.backend}".`
    );
    process.exit(1);
  }

  const adapter = new OneDriveAdapter();

  try {
    await adapter.initialize(config.connection, config.auth.credentialStore);
  } catch (err) {
    console.error(`[aifs] Adapter initialization failed: ${err.message}`);
    process.exit(1);
  }

  try {
    await startServer(adapter, config);
  } catch (err) {
    console.error(`[aifs] Server startup failed: ${err.message}`);
    process.exit(1);
  }
}

main();
