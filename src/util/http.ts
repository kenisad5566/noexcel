import axios from "axios";

/**
 * send http request
 */
export class Http {
  private config = {
    timeout: 5000,
  };

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
