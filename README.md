# OpenSheets

The open source AI agent for spreadsheets. Supports Google Sheets and Excel.

## Development

```bash
bun install
```

### Excel

```bash
bun excel:dev        # Start dev server on port 3000
bun excel:certs      # Install dev certificates
bun excel:sideload   # Load add-in in Excel
```

### Google Sheets

```bash
bun sheets:login     # Login to clasp
bun sheets:setup     # Create Apps Script project
bun sheets:push      # Push code to Google Sheets
```

### Build

```bash
bun run build        # Build for production
```

### Other

```bash
bun run typecheck    # Type check
bun run lint         # Lint
bun run lint:fix     # Lint and fix
```

