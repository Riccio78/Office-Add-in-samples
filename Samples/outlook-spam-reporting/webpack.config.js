/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
//const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const urlProd = "https://github.com/Riccio78/Office-Add-in-samples/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    //devtool: "source-map",
    devtool: false,
    entry: {
      // polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      // commands: "./src/commands/commands.js",
      spamreporting: ["./src/spamreporting/spamreporting.js", "./src/spamreporting/spamreporting.html"],
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
            options: {
              presets: ["@babel/preset-env"],
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
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "./src/spamreporting/spamreporting.js",
            to: "spamreporting.js",
          },
          {
            from: "./src/spamreporting/index.html",
            to: "index.html",
          },
        ],
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
      // new HtmlWebpackPlugin({
      //   filename: "commands.html",
      //   template: "./src/commands/commands.html",
      //   chunks: ["polyfill", "commands"],
      // }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      contentBase: path.join(__dirname, "dist"),
      compress: true,
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
