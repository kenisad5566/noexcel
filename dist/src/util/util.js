"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.createRandomStr = void 0;
const createRandomStr = (len = 16) => {
    let s = [];
    let hexDigits = "0123456789abcdef";
    for (let i = 0; i < len; i++) {
        s[i] = hexDigits.substr(Math.floor(Math.random() * 0x10), 1);
    }
    return s.join("");
};
exports.createRandomStr = createRandomStr;
//# sourceMappingURL=util.js.map