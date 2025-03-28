/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://brave-mushroom-07685f203.6.azurestaticapps.net/";

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
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.jsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".js", ".jsx", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: {
            loader: "babel-loader",
          },
          exclude: /node_modules/,
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      // Genereer taskpane.html in de output-map op basis van src/taskpane/taskpane.html
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),

      // Kopieer assets, manifestbestanden, index.html en fallbackauthdialog.html
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                // Vervang alle voorkomens van de dev-URL door de productie-URL
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            // Kopieer fallbackauthdialog.html vanuit src/public
            from: "src/public/fallbackauthdialog.html",
            to: "fallbackauthdialog.html",
          },
          {
            // Kopieer index.html vanuit de public-map naar de root van de output-map
            from: "src/public/index.html",
            to: "index.html",
          },
        ],
      }),

      // Genereer commands.html in de output-map op basis van src/commands/commands.html
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),

      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
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
