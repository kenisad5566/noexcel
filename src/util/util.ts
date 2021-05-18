export const createRandomStr = (len: number = 16): string => {
  let s: any = [];
  let hexDigits = "0123456789abcdef";
  for (let i = 0; i < len; i++) {
    s[i] = hexDigits.substr(Math.floor(Math.random() * 0x10), 1);
  }
  return s.join("");
};
