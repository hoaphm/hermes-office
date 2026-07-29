/* global require, module, process, __dirname */

const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// Caddy serves this add-in's dist under /word/ (see the repo-root Caddyfile),
// so the production base must carry that path segment.
const urlProd = "https://localhost:8643/word/";
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
      // entirely in manifest.xml, so there was no runtime code to load and
      // the manifest never referenced the emitted chunk.
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
          use: { loader: "babel-loader" },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          // The taskpane HTML links assets/design-system.css, but the
          // shared CSS is only available in the dist output (copied by
          // CopyWebpackPlugin from ../shared/design-system.css) — there's
          // no source at src/taskpane/assets/. Disable html-loader's
          // resource resolution so the <link> is left alone and shipped
          // to the dist as a plain href.
          use: {
            loader: "html-loader",
            options: { sources: false },
          },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          // Shared design system — used by both add-ins. Copied verbatim so
          // the dist HTML can <link rel="stylesheet" href="assets/design-system.css">
          // without any CSS loader pipeline.
          {
            from: "../shared/design-system.css",
            to: "assets/design-system.css",
          },
          {
            from: "manifest*.xml",
            to: "[name][ext][query]",
            transform(content) {
              // In dev the manifest is already pointing at the webpack
              // dev-server on :3000, so leave it untouched.
              if (dev) return content;
              return content.toString().replace(urlDevPattern, urlProd);
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
  };

  if (dev) {
    config.devServer = {
      // Serve the repo root, not word/, so the task pane's fetch("/config.json")
      // resolves to the same config.json Caddy serves in production. Pointing
      // this at process.cwd() (= word/) meant dev could never load a provider.
      static: [
        { directory: path.resolve(__dirname, ".."), publicPath: "/" },
        { directory: path.resolve(__dirname, "dist"), publicPath: "/" },
      ],
      https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
      // No /v1 proxy: add-ins call the configured provider directly from the
      // WebView. The old proxy pointed at a Caddy reverse-proxy hop that no
      // longer exists.
    };
  }

  return config;
};
