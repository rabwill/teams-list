'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

//make teams react components work with webpack, install file-loader and add the rule
const node= {
    net: 'empty',
    tls: 'empty',
    fs:'empty'
  };
const loaderRule = { "use": [{ "loader": "file-loader" }], "test": /\.(eot|svg|ttf|woff|woff2)$/ };
build.configureWebpack.mergeConfig({
node:node,
additionalConfiguration: (generatedConfiguration) => {
generatedConfiguration.module.rules.push(loaderRule);
return generatedConfiguration;}
});
build.initialize(gulp);