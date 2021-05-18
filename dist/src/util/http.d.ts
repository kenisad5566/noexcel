/**
 * 发送请求
 */
export declare class Http {
    private config;
    /**
     * 发送get请求
     * @param url
     */
    get(url: string, config?: {}): Promise<any>;
    post(url: string, data: any, config?: {}): Promise<any>;
}
