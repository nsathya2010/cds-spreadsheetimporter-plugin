const cds = global.cds || require('@sap/cds');
const XLSX = require('xlsx');
const path = require('path');
const SheetHandler = require('./utils/SheetHandler');
const Parser = require('./utils/Parser');

module.exports = class ImporterService extends cds.ApplicationService {
  init() {
    this.on('READ', 'Spreadsheet', async (req) => {
      const entityName = req.params?.[0]?.entity || req.data?.entity;
      if (!entityName) {
        req.error(400, 'Missing entity parameter for template download');
        return;
      }

      const entity = cds.entities()[entityName];
      if (!entity) {
        req.error(400, `Entity '${entityName}' not found`);
        return;
      }

      const templateBuffer = this._buildTemplateWorkbook(entity);
      return [{ entity: entityName, content: templateBuffer }];
    });

    this.on('UPDATE', 'Spreadsheet', async (req) => {
      try {
        console.log(
          'Spreadsheet importer received request with content type:',
          req.headers && req.headers['content-type']
        );
        console.log('Entity parameter:', req.params[0].entity);
        console.log('Request data structure:', Object.keys(req.data || {}));

        const entityName = req.params?.[0]?.entity;
        const entity = cds.entities()[entityName];
        if (!entity) {
          req.error(400, `Entity '${entityName}' not found`);
          return;
        }

        if (!req.data?.content || typeof req.data.content.on !== 'function') {
          req.error(400, 'Spreadsheet content stream is missing or invalid');
          return;
        }

        // Handle file content using a streaming approach for large files
        const chunks = [];

        // Check if we have access to the content as a stream
        console.log('Processing content as a stream');

        // Collect all chunks before processing
        await new Promise((resolve, reject) => {
          req.data.content.on('data', (chunk) => {
            console.log(`Received chunk of size: ${chunk.length} bytes`);
            chunks.push(chunk);
          });

          req.data.content.on('end', () => {
            console.log(`Stream ended, received ${chunks.length} chunks`);
            resolve();
          });

          req.data.content.on('error', (err) => {
            console.error('Error reading content stream:', err);
            reject(err);
          });
        });

        // Once we have all chunks, process the file
        const totalBuffer = Buffer.concat(chunks);
        console.log(
          `Processing complete file of size: ${totalBuffer.length} bytes`
        );

        try {
          const spreadSheet = XLSX.read(totalBuffer, {
            type: 'buffer',
            cellNF: true,
            cellDates: true,
            cellText: true,
            cellFormula: true,
          });

          let spreadsheetSheetsData = [];
          let columnNames = [];

          console.log(
            `Workbook contains ${spreadSheet.SheetNames.length} sheets`
          );

          // Loop over the sheet names in the workbook
          for (const sheetName of Object.keys(spreadSheet.Sheets)) {
            console.log(`Processing sheet: ${sheetName}`);
            let currSheetData = SheetHandler.sheet_to_json(
              spreadSheet.Sheets[sheetName]
            );
            console.log(
              `Sheet ${sheetName} has ${currSheetData.length} rows of data`
            );

            for (const dataVal of currSheetData) {
              Object.keys(dataVal).forEach((key) => {
                dataVal[key].sheetName = sheetName;
              });
            }

            spreadsheetSheetsData = spreadsheetSheetsData.concat(currSheetData);
            columnNames = columnNames.concat(
              XLSX.utils.sheet_to_json(spreadSheet.Sheets[sheetName], {
                header: 1,
              })[0]
            );
          }

          console.log(
            `Total data rows to process: ${spreadsheetSheetsData.length}`
          );
          const data = Parser.parseSpreadsheetData(
            spreadsheetSheetsData,
            entity.elements
          );

          const postProcessResult = await this._runPostProcessor({
            req,
            entity,
            data,
            workbook: {
              sheetNames: spreadSheet.SheetNames,
            },
          });

          if (postProcessResult?.runDefaultInsert === true) {
            console.log(
              `Post processor requested default insert for ${data.length} rows into ${entity.name}`
            );
            await cds.db.run(INSERT(data).into(entity.name));
          } else if (!postProcessResult) {
            console.log(`Inserting ${data.length} rows into ${entity.name}`);
            await cds.db.run(INSERT(data).into(entity.name));
          } else {
            console.log('Post processing completed without default insert');
          }

          console.log('Import completed successfully');
          return (
            postProcessResult?.response || {
              entity: entity.name,
              rows: data.length,
              inserted:
                !postProcessResult || postProcessResult.runDefaultInsert === true,
            }
          );
        } catch (xlsxError) {
          console.error('Error processing Excel file:', xlsxError);
          req.error(400, `Failed to parse spreadsheet: ${xlsxError.message}`);
          return;
        }
      } catch (error) {
        console.error('Spreadsheet import error:', error);
        req.error(500, `Failed to process spreadsheet: ${error.message}`);
      }
    });
    return super.init();
  }

  _buildTemplateWorkbook(entity) {
    const templateColumns = this._getTemplateColumns(entity);
    const headers = templateColumns.map((column) => column.name);
    const sampleRow = templateColumns.map((column) =>
      this._getSampleValueForElement(column.element)
    );

    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([headers, sampleRow]);
    XLSX.utils.book_append_sheet(workbook, sheet, 'Template');

    return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  }

  _getTemplateColumns(entity) {
    const entries = Object.entries(entity.elements || {});
    return entries
      .filter(([, element]) => {
        return (
          !element.isAssociation &&
          !element.isComposition &&
          element.virtual !== true
        );
      })
      .map(([name, element]) => ({ name, element }));
  }

  _getSampleValueForElement(element) {
    switch (element.type) {
      case 'cds.Boolean':
        return true;
      case 'cds.Date':
        return '2026-01-01';
      case 'cds.DateTime':
      case 'cds.DateTimeOffset':
        return '2026-01-01T12:00:00Z';
      case 'cds.Time':
      case 'cds.TimeOfDay':
        return '12:00:00';
      case 'cds.UInt8':
      case 'cds.Int16':
      case 'cds.Int32':
      case 'cds.Integer':
      case 'cds.Int64':
      case 'cds.Integer64':
      case 'cds.Byte':
      case 'cds.SByte':
        return 1;
      case 'cds.Double':
      case 'cds.Decimal':
        return 1.23;
      default:
        return '';
    }
  }

  async _runPostProcessor(context) {
    const pluginConfig =
      cds.env?.spreadsheetimporter ||
      cds.env?.requires?.['cds-spreadsheetimporter-plugin'] ||
      {};

    const processorPath =
      pluginConfig.postProcessor || pluginConfig.postProcessorModule;
    if (!processorPath) {
      return null;
    }

    const resolvedPath = path.isAbsolute(processorPath)
      ? processorPath
      : path.resolve(cds.root, processorPath);

    let postProcessor = require(resolvedPath);
    if (postProcessor && typeof postProcessor.process === 'function') {
      postProcessor = postProcessor.process;
    }

    if (typeof postProcessor !== 'function') {
      throw new Error(
        `Configured post processor '${resolvedPath}' does not export a function`
      );
    }

    return await postProcessor(context);
  }
};
