export interface ProcessResult {
  success: boolean;
  blob?: Blob;
  fileName?: string;
  error?: string;
  logs: string[];
}

export enum ProcessingStatus {
  IDLE = 'IDLE',
  PROCESSING = 'PROCESSING',
  SUCCESS = 'SUCCESS',
  ERROR = 'ERROR'
}

export enum HeaderType {
  NONE = 'NONE',
  SCHOOL = 'SCHOOL',
  PARTY = 'PARTY',
  DEPARTMENT = 'DEPARTMENT'
}

export interface DocxOptions {
  headerType: HeaderType;
  departmentName?: string;
  removeNumbering: boolean;
  margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
  };
  font: {
    family: string;
    sizeNormal: number;
    sizeTable: number;
  };
  paragraph: {
    lineSpacing: number;
    after: number;
    indent: number;
  };
  table: {
    rowHeight: number;
  };
}