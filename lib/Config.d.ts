import { IAuthOptions } from 'node-sp-auth';
export declare class Config {
    static get(configFilePath?: string): ICrawlOptions;
}
export interface ICrawlOptions {
    auth: IAuthOptions | null;
    sites: string[];
}
