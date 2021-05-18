import axios from "axios";

/**
 * 发送请求
 */
export class Http {
  private config = {
    timeout: 5000,
  };

  /**
   * 发送get请求
   * @param url
   */
  async get(url: string, config = {}) {
    try {
      return await axios.get(url, { ...this.config, ...config });
    } catch (error) {
      return error;
    }
  }

  async post(url: string, data: any, config = {}) {
    try {
      return await axios.post(url, data, { ...this.config, ...config });
    } catch (error) {
      return error;
    }
  }
}
