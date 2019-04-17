import { ICrawlOptions } from './Config';
export declare class News {
    static crawl(config: ICrawlOptions): Promise<void>;
    static create(config: ICrawlOptions): Promise<void>;
}
