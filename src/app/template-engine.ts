import * as fs from 'fs';
import * as moment from 'moment';
import { CellRange } from './cell-range';
import { TemplateExpression, TemplatePipe } from './template-expression';
import { WorkSheetHelper } from './worksheet-helper';
import { Cell } from "exceljs";

export class TemplateEngine {
  private readonly regExpBlocks: RegExp = /\[\[.+?]]/g;
  private readonly regExpValues: RegExp = /{{.+?}}/g;

  constructor(private wsh: WorkSheetHelper, private data: any) {
  }

  public execute(): void {
    this.processBlocks(this.wsh.getSheetDimension(), this.data);
    this.processValues(this.wsh.getSheetDimension(), this.data);
  }

  private processBlocks(cellRange: CellRange, data: any): CellRange {
    if (!cellRange.valid) {
      console.log(
        'xlsx-template-ex: Process blocks failed.',
        'The cell range is invalid and will be skipped:',
        this.wsh.sheetName, cellRange
      );
      return cellRange;
    }
    let restart;
    do {
      restart = false;
      this.wsh.eachCell(cellRange, (cell: Cell) => {
        let cVal = cell.value;
        if (typeof cVal !== "string") {
          return null;
        }
        const matches = (cVal as string).match(this.regExpBlocks);
        if (!Array.isArray(matches) || !matches.length) {
          return null;
        }

        matches.forEach((rawExpression: string) => {
          const tplExp = new TemplateExpression(rawExpression, rawExpression.slice(2, -2));
          cVal = (cVal as string).replace(tplExp.rawExpression, '');
          cell.value = cVal;
          cellRange = this.processBlockPipes(cellRange, cell, tplExp.pipes, data[tplExp.valueName]);
        });

        restart = true;
        return false;
      });
    } while (restart);
    return cellRange;
  }

  private processValues(cellRange: CellRange, data: any): void {
    if (!cellRange.valid) {
      console.log(
        'xlsx-template-ex: Process values failed.',
        'The cell range is invalid and will be skipped:',
        this.wsh.sheetName, cellRange
      );
      return;
    }
    this.wsh.eachCell(cellRange, (cell: Cell) => {
      let cVal = cell.value;
      if (typeof cVal !== "string") {
        return;
      }
      const matches = cVal.match(this.regExpValues);
      if (!Array.isArray(matches) || !matches.length) {
        return;
      }

      matches.forEach((rawExpression: string) => {
        const tplExp = new TemplateExpression(rawExpression, rawExpression.slice(2, -2));
        let resultValue: any = data[tplExp.valueName] || '';
        resultValue = this.processValuePipes(cell, tplExp.pipes, resultValue);
        cVal = (cVal as string).replace(tplExp.rawExpression, resultValue);
      });
      cell.value = cVal;
    });
  }

  private processValuePipes(cell: Cell, pipes: TemplatePipe[], value: any): string {
    try {
      pipes.forEach((pipe: TemplatePipe) => {
        switch (pipe.pipeName) {
          case 'date':
            // value = this.valuePipeDate(value, ...pipe.pipeParameters);
            value = this.valuePipeDate(value);
            break;
          case 'image':
            // value = this.valuePipeImage(cell, value, ...pipe.pipeParameters);
            value = this.valuePipeImage(cell, value);
            break;
          case 'find':
            value = this.valuePipeFind(value, ...pipe.pipeParameters);
            break;
          case 'get':
            value = this.valuePipeGet(value, ...pipe.pipeParameters);
            break;
          case 'time':
            value = this.valuePipeTime(value);
            break;
          case 'datetime':
            value = this.valuePipeDateTime(value);
            break;
          default:
            value = 'xlsx-template-ex: The value pipe not found:' + pipe.pipeName;
            console.warn(value);
        }
      });
    } catch (error) {
      console.error('xlsx-template-ex: Error on process values of pipes', error);
      return 'xlsx-template-ex: Error on process values of pipes. Look for more details in a console.';
    }
    return value || '';
  }

  private processBlockPipes(cellRange: CellRange, cell: Cell, pipes: TemplatePipe[], data: any): CellRange {
    // console.log('bp', pipes, data);
    const newRange = CellRange.createFromRange(cellRange);
    let insertedRows;
    try {
      pipes.forEach((pipe: TemplatePipe) => {
        switch (pipe.pipeName) {
          case 'repeat-rows':
            // insertedRows = this.blockPipeRepeatRows.apply(this, [cell, data].concat(pipe.pipeParameters));
            insertedRows = this.blockPipeRepeatRows(cell, data, ...pipe.pipeParameters);
            newRange.bottom += insertedRows;
            break;
          case 'tile':
            insertedRows = this.blockPipeTile(cell, data, ...pipe.pipeParameters);
            newRange.bottom += insertedRows;
            break;
          case 'filter':
            data = this.blockPipeFilter(data, ...pipe.pipeParameters);
            break;
          default:
            console.warn('xlsx-template-ex: The block pipe not found:', pipe.pipeName, pipe.pipeParameters);
        }
      });
    } catch (error) {
      console.error('xlsx-template-ex: Error on process a block of pipes', error);
      cell.value = 'xlsx-template-ex: Error on process a block of pipes. Look for more details in a console.';
    }
    return newRange;
  }

  private valuePipeDate(date?: number | string): string {
    return date ? moment(new Date(date)).format('DD.MM.YYYY') : '';
  }


  private valuePipeTime(date?: number | string): string {
    return date ? moment(new Date(date)).format('hh:mm:ss') : '';
  }

  private valuePipeDateTime(date?: number | string): string {
    return date ? moment(new Date(date)).format('DD.MM.YYYY hh:mm:ss') : '';
  }

  private valuePipeImage(cell: Cell, fileName: string): string {
    if (fs.existsSync(fileName)) {
      this.wsh.addImage(fileName, cell);
      return fileName;
    }
    return `File "${fileName}" not found`;
  }

  /** Find object in array by value of a property */
  private valuePipeFind(arrayData: any[], propertyName?: string, propertyValue?: string): any | null {
    if (Array.isArray(arrayData) && propertyName && propertyName) {
      return arrayData.find(item => item && item[propertyName] == propertyValue);
    }
    return null;
  }

  private valuePipeGet(data: any[], propertyName?: string): any | null {
    return data && propertyName && data[propertyName] || null;
  }

  private blockPipeFilter(dataArray: any[], propertyName?: string, propertyValue?: string): any[] {
    if (Array.isArray(dataArray) && propertyName) {
      if (propertyValue) {
        return dataArray.filter(item => typeof item === "object" && item[propertyName] === propertyValue);
      }
      return dataArray.filter(item => typeof item === "object" &&
        item.hasOwnProperty(propertyName) && item[propertyName]
      );
    }
    return dataArray;
  }

  /** @return {number} count of inserted rows */
  blockPipeRepeatRows(cell: Cell, dataArray: any[], countRows?: number | string): number {
    if (!Array.isArray(dataArray) || !dataArray.length) {
      console.warn('TemplateEngine.blockPipeRepeatRows', cell.address,
        'The data must be not empty array, but got:', dataArray
      );
      return 0;
    }
    countRows = +countRows > 0 ? +countRows : 1;
    const startRow = +cell.row;
    const endRow = startRow + countRows - 1;
    if (dataArray.length > 1) {
      this.wsh.cloneRows(startRow, endRow, dataArray.length - 1);
    }

    const wsDimension = this.wsh.getSheetDimension();
    let sectionRange = new CellRange(startRow, wsDimension.left, endRow, wsDimension.right);

    dataArray.forEach(data => {
      sectionRange = this.processBlocks(sectionRange, data);
      this.processValues(sectionRange, data);
      sectionRange.move(+countRows, 0);
    });
    return (dataArray.length - 1) * countRows;
  }

  /** @return {number} count of inserted rows */
  private blockPipeTile(cell: Cell, dataArray: any[], blockRows?: number | string, blockColumns?: number | string,
    tileColumns?: number | string): number {
    // return;
    if (!Array.isArray(dataArray) || !dataArray.length) {
      console.warn('TemplateEngine.blockPipeTile', cell.address,
        'The data must be not empty array, but got:', dataArray
      );
      return 0;
    }
    blockRows = +blockRows > 0 ? +blockRows : 1;
    blockColumns = +blockColumns > 0 ? +blockColumns : 1;
    tileColumns = +tileColumns > 0 ? +tileColumns : 1;

    const blockRange = new CellRange(
      +cell.row, +cell.col, +cell.row + blockRows - 1, +cell.col + blockColumns - 1
    );
    const cloneRowsCount = Math.ceil(dataArray.length / tileColumns) - 1;
    if (dataArray.length > tileColumns) {
      this.wsh.cloneRows(blockRange.top, blockRange.bottom, cloneRowsCount);
    }

    let tileColumn = 1, tileRange = CellRange.createFromRange(blockRange);
    dataArray.forEach((data, idx: number, array: any[]) => {
      // Prepare the next tile
      if ((idx !== array.length - 1) && (tileColumn + 1 <= tileColumns)) {
        const nextTileRange = CellRange.createFromRange(tileRange);
        nextTileRange.move(0, tileRange.countColumns);
        this.wsh.copyCellRange(tileRange, nextTileRange);
      }

      // Process templates
      tileRange = this.processBlocks(tileRange, data);
      this.processValues(tileRange, data);

      // Move tiles
      if (idx !== array.length - 1) {
        tileColumn++;
        if (tileColumn <= tileColumns) {
          tileRange.move(0, tileRange.countColumns);
        } else {
          tileColumn = 1;
          blockRange.move(tileRange.countRows, 0);
          tileRange = CellRange.createFromRange(blockRange);
        }
      }
    });

    return cloneRowsCount * blockRange.countRows;
  }
}
