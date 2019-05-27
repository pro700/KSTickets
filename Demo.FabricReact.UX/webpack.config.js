var path = require('path');
const SPSaveWebpackPlugin = require('spsave-webpack-plugin');

module.exports = {

    //externals: [{
    //    "jQuery": "jQuery",
    //    "sp-init": "sp-init",
    //    "microsoft-ajax": "microsoft-ajax",
    //    "sp-runtime": "sp-runtime",
    //    "sharepoint": "sharepoint"
    //}],

    // Target the output of the typescript compiler
    context: path.join(__dirname, "src"),

    // File(s) to target in the 'build' directory
    entry: './index.tsx',

    // Output
    output: {
        filename: 'bundle.js',
        path: path.resolve(__dirname, 'dist'),
        publicPath: "/dist/"
    },

    plugins: [new SPSaveWebpackPlugin({
        coreOptions: {
            checkin: true,
            checkinType: 1,
            siteUrl: "http://tickets"
        },
        credentialOptions: {
            /* See https://github.com/s-KaiNet/node-sp-auth#params for authentication options */
            username: 'sergiy.sokol',
            password: 'Sok%asw8',
            domain: 'KYIVSTAR.UA'
        },
        fileOptions: {
            folder: "SiteAssets/KS/CEWP/Tickets",
            fileName: 'bundle.js'
        }
    })],

    // Resolve the file extensions
    resolve: {
        extensions: [".js", ".jsx", ".ts", ".tsx"]
    },

    // Module to define what libraries with the compiler
    module: {
        // Loaders
        loaders: [
            {
                // Target the sass files
                test: /\.scss?$/,

                // Define the compiler to use
                use: [
                    // Create style nodes from the CommonJS code
                    { loader: "style-loader" },
                    // Translate css to CommonJS
                    { loader: "css-loader" },
                    // Compile sass to css
                    { loader: "sass-loader" }
                ]
            },
            {
                // Target the .ts and .tsx files
                test: /\.tsx$/,
                // Exclude the node modules folder
                exclude: /node_modules/,
                // Define the compiler to use
                use: [
                    {
                        // Compile the JSX code to javascript
                        loader: "babel-loader",
                        // Options
                        options: {
                            // Ensure the javascript works in legacy browsers
                            presets: ["es2015"]
                        }
                    },
                    {
                        // Compile the typescript code to JSX
                        loader: "ts-loader"
                    }
                ]
            }
        ]
    }
};