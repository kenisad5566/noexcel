/**
 * send http request
 */
export declare class Http {
    private config;
    get(url: string, config?: {}): Promise<any>;
    post(url: string, data: any, config?: {}): Promise<any>;
}
