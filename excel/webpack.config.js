/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlProd = "https://localhost:8643/excel/";
// String.replace() with a string pattern only substitutes the FIRST match —
// the manifest has a dozen dev URLs, so this must be a global regex. Matches
// with or without a trailing slash: resource URLs carry a path, but the
// <AppDomain> trust entry is a bare `https://localhost:3000`, and a pattern
// that demanded the slash let that one survive into production, widening
// AppDomain trust. The bare form is rewritten to the Gateway origin (no path);
// the pathed form to the add-in's base.
const urlDevPattern = /https:\/\/localhost:3000\/?/g;
const urlProdOrigin = "https://localhost:8643";

/* global require, module, process */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
      functions: "./src/functions/functions.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          // See word/webpack.config.js — we don't want html-loader to try
          // resolving assets/* from source (it lives only in the dist).
          use: {
            loader: "html-loader",
            options: { sources: false },
          },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "functions", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "../shared/design-system.css",
            to: "assets/design-system.css",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content
                  .toString()
                  // The source manifest lists BOTH AppDomains (dev :3000 and the
                  // Gateway origin :8643, which is the same in prod). Rewriting
                  // the dev one to the origin would duplicate the existing entry,
                  // so drop the dev-only line instead.
                  .replace(/\n\s*<AppDomain>\s*https:\/\/localhost:3000\s*<\/AppDomain>/g, "")
                  .replace(urlDevPattern, (m) => (m.endsWith("/") ? urlProd : urlProdOrigin));
              }
            },
          },
        ],
      }),
    ],
    devServer: {
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: [
        {
          context: ["/v1"],
          target: "https://localhost:8643",
          secure: false,
        },
      ],
    },
  };

  return config;
};
