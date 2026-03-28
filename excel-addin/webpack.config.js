const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const fs = require("fs");
const os = require("os");

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");

module.exports = {
  entry: { taskpane: "./src/taskpane.js", auth: "./src/auth.js" },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  module: {
    rules: [
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane.html",
      chunks: ["taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "manifest.xml", to: "manifest.xml" },
        { from: "assets",       to: "assets" },
      ],
    }),
  ],
  devServer: {
    port: 3000,
    hot: true,
    server: {
      type: "https",
      options: {
        key: fs.readFileSync(path.join(certDir, "localhost.key")),
        cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
      },
    },
    headers: { "Access-Control-Allow-Origin": "*" },
  },
};
