"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Http = void 0;
const tslib_1 = require("tslib");
const axios_1 = tslib_1.__importDefault(require("axios"));
/**
 * 发送请求
 */
class Http {
    constructor() {
        this.config = {
            timeout: 5000,
        };
    }
    /**
     * 发送get请求
     * @param url
     */
    async get(url, config = {}) {
        try {
            return await axios_1.default.get(url, Object.assign(Object.assign({}, this.config), config));
        }
        catch (error) {
            return error;
        }
    }
    async post(url, data, config = {}) {
        try {
            return await axios_1.default.post(url, data, Object.assign(Object.assign({}, this.config), config));
        }
        catch (error) {
            return error;
        }
    }
}
exports.Http = Http;
//# sourceMappingURL=http.js.map