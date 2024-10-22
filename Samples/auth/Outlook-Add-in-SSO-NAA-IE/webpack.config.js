/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

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
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      dialoginternetexplorer: ["./src/taskpane/fallback/fallbackauthdialoginternetexplorer.ts"],
      dialog: ["./src/taskpane/fallback/fallbackauthdialog.ts"],
      signoutdialoginternetexplorer: ["./src/taskpane/fallback/signoutdialoginternetexplorer.ts"],
      signoutdialog: ["./src/taskpane/fallback/signoutdialog.ts"],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
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
      new HtmlWebpackPlugin({
        filename: "auth.html",
        template: "./src/taskpane/fallback/auth.html",
        chunks: [],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/taskpane/fallback/dialog.html",
        chunks: ["dialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "signoutdialoginternetexplorer.html",
        template: "./src/taskpane/fallback/dialog.html",
        chunks: ["polyfill", "signoutdialoginternetexplorer"],
      }),
      new HtmlWebpackPlugin({
        filename: "signoutdialog.html",
        template: "./src/taskpane/fallback/dialog.html",
        chunks: ["polyfill", "signoutdialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialoginternetexplorer.html",
        template: "./src/taskpane/fallback/dialog.html",
        chunks: ["polyfill", "dialoginternetexplorer"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
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
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
