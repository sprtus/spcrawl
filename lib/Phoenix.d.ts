import { ICrawlOptions } from './Config';
export declare class Phoenix {
    static crawl(config: ICrawlOptions): Promise<void>;
    static create(config: ICrawlOptions): Promise<void>;
}
