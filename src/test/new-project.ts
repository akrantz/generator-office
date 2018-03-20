/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as helpers from 'yeoman-test';
import * as assert from 'yeoman-assert';
import * as path from 'path';

/**
 * Test addin from user answers
 * new project, default folder, defaul host.
 */
describe('new project - answers', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: projectDisplayName,
    host: 'excel',
    isManifestOnly: false,
    ts: null,
    framework: null,
    open: false
  };
  let manifestFileName = projectEscapedName + '-manifest.xml';

  /** Test addin when user chooses jquery and typescript. */
  describe('jquery & typescript', () => {
    before((done) => {
      answers.ts = true;
      answers.framework = 'jquery';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses jquery and javascript. */
  describe('jquery & javascript', () => {
    before((done) => {
      answers.ts = false;
      answers.framework = 'jquery';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'app.css',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png',
        'function-file/function-file.html',
        'function-file/function-file.js',
        'bsconfig.json',
        'app.js',
        'index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses angular and typescript. */
  describe('angular & typescript', () => {
    before((done) => {
      answers.ts = true;
      answers.framework = 'angular';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses angular and javascript. */
  describe('angular & javascript', () => {
    before((done) => {
      answers.ts = false;
      answers.framework = 'angular';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'app.css',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png',
        'function-file/function-file.html',
        'function-file/function-file.js',
        'bsconfig.json',
        'app.js',
        'index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user chooses react and typescript. */
  describe('react & typescript', () => {
    before((done) => {
      answers.ts = true;
      answers.framework = 'react';
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'config/webpack.common.js',
        'config/webpack.dev.js',
        'config/webpack.prod.js',
        'src/assets/styles/global.less',
        'src/components/App.tsx',
        'src/components/Header.tsx',
        'src/components/HeroList.tsx',
        'src/components/Progress.tsx',
        'src/index.html',
        'src/index.tsx',
        'tslint.json',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });
});

/**
 * Test addin from user answers and arguments
 * new project, default folder, typescript, jquery.
 */
describe('new project - answers & args - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: null,
    host: null,
    isManifestOnly: false,
    ts: true,
    framework: null,
    open: false
  };
  let argument = [];

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in"
	 */
  describe('command line: --name "Display Name"', () => {
    before((done) => {
      answers.host = 'excel';
      answers.framework = 'jquery';

      helpers.run(path.join(__dirname, '../app'))
        .withArguments([])
        .withOptions({ name: projectDisplayName })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let manifestFileName = projectEscapedName + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel"
	 */
  describe('command line: --name "Display name" --host excel', () => {
    before((done) => {
      answers.framework = 'jquery';

      helpers.run(path.join(__dirname, '../app'))
        .withArguments([])
        .withOptions({ name: projectDisplayName }, { host: 'excel' })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let manifestFileName = projectEscapedName + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel jquery"
	 */
  describe('command line: jquery --name "Display Name" --host excel', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app'))
        .withArguments(['jquery'])
        .withOptions({ name: projectDisplayName }, { host: 'excel'})
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let manifestFileName = projectEscapedName + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });
});

/**
 * Test addin from user answers and options
 * new project, default folder, typescript, jquery.
 */
describe('new project - answers & opts - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: projectDisplayName,
    host: 'excel',
    isManifestOnly: false,
    ts: null,
    framework: 'jquery',
    open: false
  };

  let manifestFileName = projectEscapedName + '-manifest.xml';

  /** Test addin when user pass in --js. */
  describe('options: --js', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app'))
        .withOptions({ js: true })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'app.css',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png',
        'function-file/function-file.html',
        'function-file/function-file.js',
        'app.js',
        'index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });

  /** Test addin when user pass in --skip-install. */
  describe('options: --skip-install', () => {
    before((done) => {
      answers.ts = true;
      helpers.run(path.join(__dirname, '../app'))
        .withOptions({ 'skip-install': true })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'bsconfig.json',
        'src/app.css',
        'src/assets/icon-16.png',
        'src/assets/icon-32.png',
        'src/assets/icon-80.png',
        'src/assets/logo-filled.png',
        'src/function-file/function-file.html',
        'src/function-file/function-file.ts',
        'tsconfig.json',
        'src/app.ts',
        'src/index.html',
        'resource.html'
      ];

      assert.file(expected);
      done();
    });
  });
});
