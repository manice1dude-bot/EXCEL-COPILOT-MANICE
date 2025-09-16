const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');

const isProd = process.env.NODE_ENV === 'production';

module.exports = {
  mode: isProd ? 'production' : 'development',
  devtool: isProd ? 'source-map' : 'eval-source-map',
  
  entry: {
    'taskpane': './src/taskpane/index.tsx',
    'functions': './src/functions/functions.ts',
    'commands': './src/commands/ribbonCommands.ts'
  },

  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].[contenthash].js',
    publicPath: '/',
    clean: true
  },

  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx'],
    alias: {
      '@': path.resolve(__dirname, 'src'),
      '@components': path.resolve(__dirname, 'src/components'),
      '@services': path.resolve(__dirname, 'src/services'),
      '@utils': path.resolve(__dirname, 'src/utils')
    }
  },

  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: [
          {
            loader: 'ts-loader',
            options: {
              transpileOnly: true,
              experimentalWatchApi: true
            }
          }
        ],
        exclude: /node_modules/
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      },
      {
        test: /\.(png|jpg|jpeg|gif|svg)$/,
        type: 'asset/resource',
        generator: {
          filename: 'assets/images/[name][ext]'
        }
      },
      {
        test: /\.(woff|woff2|eot|ttf|otf)$/,
        type: 'asset/resource',
        generator: {
          filename: 'assets/fonts/[name][ext]'
        }
      }
    ]
  },

  plugins: [
    new CleanWebpackPlugin(),
    
    // Taskpane HTML
    new HtmlWebpackPlugin({
      template: './src/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane'],
      inject: 'body',
      minify: isProd ? {
        removeComments: true,
        collapseWhitespace: true,
        removeRedundantAttributes: true,
        useShortDoctype: true,
        removeEmptyAttributes: true,
        removeStyleLinkTypeAttributes: true,
        keepClosingSlash: true,
        minifyJS: true,
        minifyCSS: true,
        minifyURLs: true
      } : false
    }),

    // Commands HTML
    new HtmlWebpackPlugin({
      template: './src/commands.html',
      filename: 'commands.html',
      chunks: ['commands'],
      inject: 'body',
      minify: isProd ? {
        removeComments: true,
        collapseWhitespace: true,
        removeRedundantAttributes: true,
        useShortDoctype: true,
        removeEmptyAttributes: true,
        removeStyleLinkTypeAttributes: true,
        keepClosingSlash: true,
        minifyJS: true,
        minifyCSS: true,
        minifyURLs: true
      } : false
    }),

    // Functions HTML (for custom functions)
    new HtmlWebpackPlugin({
      template: './src/functions.html',
      filename: 'functions.html',
      chunks: ['functions'],
      inject: 'body',
      minify: isProd ? {
        removeComments: true,
        collapseWhitespace: true
      } : false
    }),

    // Copy static assets
    new CopyWebpackPlugin({
      patterns: [
        {
          from: './manifest.xml',
          to: 'manifest.xml',
          transform(content, path) {
            // In production, update manifest URLs to production URLs
            if (isProd) {
              return content.toString().replace(/https:\/\/localhost:3000/g, 'https://your-production-url.com');
            }
            return content;
          }
        },
        {
          from: './src/functions/functions.json',
          to: 'functions.json'
        },
        {
          from: './src/assets',
          to: 'assets',
          noErrorOnMissing: true
        }
      ]
    })
  ],

  devServer: {
    port: 3000,
    hot: true,
    open: false,
    historyApiFallback: true,
    allowedHosts: 'all',
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
      'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
    },
    https: {
      cert: './certs/server.crt',
      key: './certs/server.key',
      ca: './certs/ca.crt'
    },
    static: {
      directory: path.join(__dirname, 'dist'),
      publicPath: '/'
    },
    client: {
      overlay: {
        errors: true,
        warnings: false
      }
    }
  },

  optimization: {
    splitChunks: {
      chunks: 'all',
      cacheGroups: {
        vendor: {
          test: /[\\/]node_modules[\\/]/,
          name: 'vendors',
          chunks: 'all',
          priority: 10,
          reuseExistingChunk: true
        },
        office: {
          test: /[\\/]node_modules[\\/]@microsoft[\\/]office/,
          name: 'office-js',
          chunks: 'all',
          priority: 20
        },
        fluentui: {
          test: /[\\/]node_modules[\\/]@fluentui[\\/]/,
          name: 'fluentui',
          chunks: 'all',
          priority: 15
        }
      }
    },
    runtimeChunk: 'single'
  },

  performance: {
    maxAssetSize: 500000,
    maxEntrypointSize: 500000,
    hints: isProd ? 'warning' : false
  },

  externals: {
    // Office.js is loaded externally via CDN
    'office': 'Office'
  }
};

// Development-specific configuration
if (!isProd) {
  module.exports.optimization.minimize = false;
  module.exports.devtool = 'eval-source-map';
}

// Production-specific configuration
if (isProd) {
  const TerserPlugin = require('terser-webpack-plugin');
  const CssMinimizerPlugin = require('css-minimizer-webpack-plugin');

  module.exports.optimization.minimizer = [
    new TerserPlugin({
      terserOptions: {
        compress: {
          drop_console: true,
          drop_debugger: true
        }
      }
    }),
    new CssMinimizerPlugin()
  ];

  // Add bundle analyzer for production builds
  if (process.env.ANALYZE) {
    const { BundleAnalyzerPlugin } = require('webpack-bundle-analyzer');
    module.exports.plugins.push(
      new BundleAnalyzerPlugin({
        analyzerMode: 'static',
        openAnalyzer: false,
        reportFilename: 'bundle-report.html'
      })
    );
  }
}

module.exports.stats = {
  errorDetails: true,
  colors: true,
  modules: false,
  chunks: false,
  chunkModules: false
};