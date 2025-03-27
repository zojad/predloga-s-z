/* eslint-disable no-undef */
const path = require('path');
const devCerts = require('office-addin-dev-certs');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');

const urlDev = 'https://localhost:3001/';
const urlProd = 'https://localhost:3001/';

async function getHttpsOptions() {
  return await devCerts.getHttpsServerOptions();
}

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const httpsOptions = await getHttpsOptions();

  return {
    devtool: dev ? 'source-map' : undefined,
    entry: {
      polyfill: ['core-js/stable', 'regenerator-runtime/runtime'],
      taskpane: './src/taskpane/taskpane.js',
      commands: './src/commands/commands.js'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      clean: true,
      filename: '[name].bundle.js'
    },
    resolve: {
      extensions: ['.js']
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              presets: [['@babel/preset-env', { 
                useBuiltIns: 'usage',
                corejs: 3 
              }]]
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: 'html-loader'
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext][query]'
          }
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['polyfill', 'taskpane']
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['polyfill', 'commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets/*',
            to: 'assets/[name][ext][query]'
          },
          {
            from: 'manifest*.xml',
            to: '[name][ext]',
            transform(content) {
              return dev 
                ? content 
                : content.toString().replace(new RegExp(urlDev, 'g'), urlProd);
            }
          }
        ]
      })
    ],
    devServer: {
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      server: {
        type: 'https',
        options: httpsOptions
      },
      port: 3001,
      hot: true,
      liveReload: false,
      client: {
        overlay: {
          errors: true,
          warnings: false
        }
      }
    }
  };
};