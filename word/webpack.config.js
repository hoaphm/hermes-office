const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// Caddy serves this add-in's dist under /word/ (see the repo-root Caddyfile),
// so the production base must carry that path segment.
const urlProd = "https://localhost:8643/word/";
// String.replace() with a string pattern only substitutes the FIRST match —
// the manifest has a dozen dev URLs, so this must be a global regex.
const urlDevPattern = /https:\/\/localhost:3000\//g;

module.exports = (env, options) => {
  const dev = options.mode === "development";
  return {
    devtool: "source-map",
    entry: {
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
    },
    output: {
      clean: true,
      // Content-hash JS filenames so Office's WebView can never serve a stale
      // cached bundle after a rebuild (it caches aggressively by URL).
      filename: dev ? "[name].js" : "[name].[contenthash].js",
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
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
        chunks: ["taskpane"],
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
        ],
      }),
    ],
  };
};
