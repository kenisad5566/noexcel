const path = require('path')

module.exports = {
    mode:'none',
  entry: "./src/index.ts",
    output: {
        path:path.resolve(__dirname, 'dist'),
        filename: "index.js",
    },
    resolve: {
        // Add '.ts' and '.tsx' as a resolvable extension.
        extensions: ["", ".ts", ".js"]
    },
    module: {
        rules: [
            // all files with a '.ts' or '.tsx' extension will be handled by 'ts-loader'
            { test: /\.ts?$/, use: "ts-loader" }
        ]
    }
};