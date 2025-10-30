/* eslint-disable no-undef */
const path = require("path");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // Replace with production URL

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  return {
    devtool: "source-map",
    entry: {
      // ✅ Primary taskpane bundle
      taskpane: "./src/taskpane/entry.tsx",

      // ✅ Ribbon commands script
      commands: "./src/commands/commands.ts",

      // ✅ Bulk Archive dialog window
      dialog: "./src/commands/dialog.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      publicPath: urlDev,
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.(ts|tsx|js|jsx)$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
        {
          test: /\.html$/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/i,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      // ✅ Taskpane HTML
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        inject: "body",
        scriptLoading: "defer",
        minify: false,
        cache: false,
        templateParameters: {
          officeJsUrl: "https://appsforoffice.microsoft.com/lib/1/hosted/office.js",
        },
      }),

      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        inject: "body",
        scriptLoading: "defer",
        minify: false,
        cache: false,
        templateParameters: {
          officeJsUrl: "https://appsforoffice.microsoft.com/lib/1/hosted/office.js",
        },
      }),

      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/commands/dialog.html",
        chunks: ["dialog"],
        inject: "body",
        scriptLoading: "defer",
        minify: false,
        cache: false,
        templateParameters: {
          officeJsUrl: "https://appsforoffice.microsoft.com/lib/1/hosted/office.js",
        },
      }),

      // ✅ Copy manifest + assets
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              return dev
                ? content
                : content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],

devServer: {
  port: 3000,
  hot: false,
  liveReload: true,
  historyApiFallback: true,
  static: {
    directory: path.join(__dirname, "dist"),
    watch: true,
    serveIndex: true,
  },
  headers: {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "*",
  },
  server: {
    type: "https",
    options: await getHttpsOptions(),
  },
  allowedHosts: "all",
},

  };
};
