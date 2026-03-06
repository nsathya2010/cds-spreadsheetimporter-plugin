# cds-spreadsheetimporter-plugin

This is a plugin for the [CAP](https://cap.cloud.sap/) framework that allows you to import data from spreadsheets into your CAP project.

## Features

- Upload `.xlsx` data to a target CDS entity via OData.
- Download a sample/template `.xlsx` generated from the target entity model.
- Optional post-processing hook to handle parsed rows yourself instead of default insert.

## API

### Upload spreadsheet

`PUT /odata/v4/importer/Spreadsheet(entity='<fully.qualified.Entity>')/content`

Content type:

`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

### Download template spreadsheet

`GET /odata/v4/importer/Spreadsheet(entity='<fully.qualified.Entity>')/content`

The generated template contains:

- Header row with entity element names.
- A sample data row based on CDS data types.

## Post Processing Hook

You can configure a custom processor module. When configured, the plugin parses the spreadsheet and calls your processor with the parsed rows.

Default behavior:

- No processor configured: plugin inserts rows directly (`INSERT ... INTO <entity>`).
- Processor configured: plugin does not insert unless your processor returns `runDefaultInsert: true`.

### Configuration

In your CAP project's `package.json`:

```json
{
  "cds": {
    "spreadsheetimporter": {
      "postProcessor": "./srv/spreadsheet-post-processor.js"
    }
  }
}
```

### Processor module contract

`srv/spreadsheet-post-processor.js`:

```js
module.exports = async function processSpreadsheet(context) {
  const { req, entity, data, workbook } = context;

  // Custom logic here (e.g. queue event, call external API, batch processing)
  console.log(`Received ${data.length} rows for ${entity.name}`);
  console.log(`Workbook sheets: ${workbook.sheetNames.join(", ")}`);

  return {
    // Set true to allow plugin's default INSERT after your processing.
    runDefaultInsert: false,

    // Optional custom response sent back to caller.
    response: {
      entity: entity.name,
      rows: data.length,
      inserted: false,
      message: "Rows were handled by custom post processor",
    },
  };
};
```

## Release Process

This project uses [release-it](https://github.com/release-it/release-it) to automate version management and package publishing. The release workflow is configured through GitHub Actions and can be triggered in two ways:

### Manual Release

1. Go to the GitHub repository's "Actions" tab
2. Select the "Release" workflow
3. Click "Run workflow"
4. You can either:
   - Leave the version field empty for automatic versioning based on conventional commits
   - Specify a specific version (e.g., "1.2.0")

### Local Release (for maintainers)

If you prefer to release from your local machine:

1. Ensure you have the necessary credentials:
   - `NPM_TOKEN` for publishing to npm
   - `GITHUB_TOKEN` for creating GitHub releases
2. Run one of the following commands:
   ```bash
   npm run release              # For automatic versioning
   npm run release X.Y.Z        # For specific version
   ```

### What happens during release?

The release process will:

1. Determine the next version number (based on conventional commits or manual input)
2. Update the package.json version
3. Generate/update the CHANGELOG.md file
4. Create a git tag
5. Push changes to GitHub
6. Create a GitHub release with release notes
7. Publish the package to npm

The release configuration uses the Angular conventional commit preset for changelog generation and requires commit messages to follow the conventional commits specification.
