// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2020  Mozard BV
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://officetest.mozard.nl/";

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      vendor: ["react", "react-dom", "core-js", "office-ui-fabric-react"],
      polyfill: "babel-polyfill",
      taskpane: ["react-hot-loader/patch", "./src/taskpane/index.js"],
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      alias: {
        "react-dom": "@hot-loader/react-dom",
      },
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: ["react-hot-loader/webpack", "babel-loader"],
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: "file-loader",
          options: {
            name: "[path][name].[ext]",
          },
        },
      ],
    },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: "taskpane.css",
            from: "./src/taskpane/taskpane.css",
          },
          {
            to: "[name]." + buildType + ".[ext]",
            from: "manifest*.xml",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new MiniCssExtractPlugin({
        filename: "[name].[fullhash].css",
      }),
      new HtmlWebpackPlugin({
        filename: "index.html",
        template: "./src/taskpane/index.html",
        templateParameters: {
          mode: options.mode,
        },
        chunks: ["taskpane", "vendor", "polyfill"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
        "Access-Control-Allow-Headers": "X-Requested-With, Content-Type, Authorization",
      },
      https: options.https !== undefined ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: {
        "/public": {
          target: "https://office.mozard.nl",
          secure: true,
          headers: {
            Host: "office.mozard.nl",
          },
        },
      },
      publicPath: "/",
    },
  };

  return config;
};
