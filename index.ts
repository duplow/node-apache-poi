import { WritableStream, ReadableStream } from 'stream/web';

export interface IPOI {
  Excel: IExcel,
  Word: IWord
}

export interface IColumn {
  index: number
}

export interface IRow {
  index: number
}

export enum ICellValueType {
  TEXT = 'text',
  NUMERIC = 'numeric'
}

export interface ICell {
  columnNumber: number,
  rowNumber: number,
  valueType: ICellValueType
  getValueAsText(): string
  setValue(value: string)
}

export interface ISheet {
  name: string
  index: number
  isActive: boolean
  lastColumnNumber: number
  lastRowNumber: number
  getDocument() : IExcel,
  getRowAt(index: number) : IRow
  getColumnAt(index: number) : IColumn
  getCellAt(columnNumber: number, rowNumber: number) : ICell
  setDocument(document: IExcel) : ISheet
  setActive(isActive: boolean) : ISheet
  setName(name: string) : ISheet
  setIndex(index: number) : ISheet
}

export interface IDocument {
  isOpen: boolean
  isXmlDocument: boolean
  setFile(path: string) : void
  openFile(file: string): void
  writeFile(file: string | WritableStream): void
  open(): Promise<IDocument>
}

// HSSF
export interface IExcel extends IDocument {
  activeSheetIndex: number
  getSheets(): ISheet[]
  addSheet(sheet: ISheet) : IExcel
  setActiveSheet(index: number) : IExcel
  newSheet(name: string) : ISheet
}

// XSSF
export interface IWord extends IDocument {

}

export const isXmlDocument = (document: IDocument) : boolean => {
  return document.isXmlDocument
}

export class Excel implements IExcel {
  protected _isOpen: boolean;
  protected _isXmlDocument: boolean;
  protected _activeSheetIndex: number;
  protected _filename: string;

  public get isOpen() {
    return this._isOpen;
  }

  public get isXmlDocument() {
    return this._isXmlDocument;
  }

  public get activeSheetIndex() {
    return this._activeSheetIndex;
  }

  public get filename() {
    return this._filename
  }

  async open(): Promise<IExcel> {
    if (this._isOpen) {
      throw new Error('FileAlreadyOpen')
    }

    // Load document info

    this._isOpen = true;
    this._isXmlDocument = true;
    this._activeSheetIndex = 0;
  
    return this;
  }

  getSheets(): ISheet[] {
    throw new Error('Method not implemented.');
  }

  addSheet(sheet: ISheet): IExcel {
    throw new Error('Method not implemented.');
  }

  setActiveSheet(index: number): IExcel {
    throw new Error('Method not implemented.');
  }

  newSheet(name: string): ISheet {
    throw new Error('Method not implemented.');
  }

  setFile(path: string): void {
    this._filename = path;
  }

  async openFile(path: string): Promise<void> {
    this._filename = path;
    await this.open();
  }

  async writeFile(path: string): Promise<void> {
    // TODO: Code here
  }
}