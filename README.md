# OpenSheets

The open source AI agent for spreadsheets. Supports Google Sheets and Excel.

## Development

```bash
bun install
```

### Excel

```bash
bun --filter @repo/addins excel:certs      # Install dev certificates
bun --filter @repo/addins excel:sideload   # Load add-in in Excel
bun --filter @repo/addins excel:dev        # Start dev server on port 3000
```

### Google Sheets

```bash
bun --filter @repo/addins sheets:login     # Login to clasp
bun --filter @repo/addins sheets:setup     # Create Apps Script project
bun --filter @repo/addins sheets:push      # Push code to Google Sheets
```

### Build

```bash
bun --filter @repo/addins run build        # Build for production
```

### Other

```bash
bun run typecheck    # Type check
bun run lint         # Lint
bun run lint:fix     # Lint and fix
```

