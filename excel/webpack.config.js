/* global require, module, process, __dirname */

const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlProd = "https://localhost:8643/excel/";
// String.replace() with a string pattern only substitutes the FIRST match —
// the manifest has a dozen dev URLs, so this must be a global regex.
const urlDevPattern = /https:\/\/localhost:3000\//g;

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    // Source maps are a dev aid; the production bundle is served locally and
    // does not need to ship them.
    devtool: dev ? "source-map" : false,
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      // No `commands` entry: the ribbon button uses ShowTaskpane, declared
      // entirely in manifest.xml. The chunk held nothing but an empty
      // Office.onReady() and was being injected into taskpane.html, where it
      // ran a second, pointless onReady.
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
        chunks: ["polyfill", "taskpane", "functions"],
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
                return content.toString().replace(urlDevPattern, urlProd);
              }
            },
          },
          // No config.json is copied here. The task pane fetches the
          // root-relative "/config.json", which Caddy serves from the repo
          // root, so a copy in dist/ was never read — it only duplicated the
          // API key into the folder users are told to open when sideloading,
          // and each rebuild overwrote the real key with the placeholder.
        ],
      }),
    ],
    devServer: {
      // Serve the repo root so the task pane's fetch("/config.json") resolves
      // to the same config.json Caddy serves in production.
      static: [
        { directory: path.resolve(__dirname, ".."), publicPath: "/" },
        { directory: path.resolve(__dirname, "dist"), publicPath: "/" },
      ],
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      // No /v1 proxy and no wildcard CORS header: add-ins call the configured
      // provider directly from the WebView, and everything the pane fetches
      // from this server is same-origin. The old proxy pointed at a Caddy
      // reverse-proxy hop that no longer exists.
    },
  };

  return config;
};
