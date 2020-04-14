/// <reference types="node" />
import { Transform } from "stream";
import { IXlsxStreamOptions, IWorksheetOptions } from "./types";
export declare function getXlsxStream(options: IXlsxStreamOptions): Promise<Transform>;
export declare function getWorksheets(options: IWorksheetOptions): Promise<string[]>;
