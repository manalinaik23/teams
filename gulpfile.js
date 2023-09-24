'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

const path = require('path');

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if(!generatedConfiguration.resolve.alias){
      generatedConfiguration.resolve.alias = {};
    }

    // web part specific components folder
    generatedConfiguration.resolve.alias['@common'] = path.resolve( __dirname, 'lib/common/')

    // // shared components
    // generatedConfiguration.resolve.alias['@components'] = path.resolve( __dirname, 'lib/shared/components/')

    //root src folder
    generatedConfiguration.resolve.alias['@src'] = path.resolve( __dirname, 'lib')

    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
