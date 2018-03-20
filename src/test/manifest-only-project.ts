/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as helpers from 'yeoman-test';
import * as assert from 'yeoman-assert';
import * as path from 'path';

/**
 * Test addin from user answers
 * manifest-only project, default folder, defaul host.
 */
describe('manifest-only project - answers', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: false,
    name: projectDisplayName,
    host: 'excel',
    isManifestOnly: true,
    ts: null,
    framework: null,
    open: false
  };
  let manifestFileName = projectEscapedName + '-manifest.xml';

	/** Test addin when user chooses jquery and typescript. */
  describe('manifest-only', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app'))
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let expected = [
        manifestFileName,
        'package.json',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png'
      ];

      assert.file(expected);
      done();
    });
  });
});

/**
 * Test addin from user answers and arguments
 * manifest-only project, default folder, typescript, jquery.
 */
describe('manifest-only project - answers & args - jquery & typescript', () => {
  let projectDisplayName = 'My Office Add-in';
  let projectEscapedName = 'my-office-add-in';
  let answers = {
    folder: null,
    name: null,
    host: null,
    isManifestOnly: null,
    ts: null,
    framework: null,
    open: false
  };

	/**
	 * Test addin when user provides --name option
	 * "my-office-add-in"
	 */
  describe('command line: --name "Display Name"', () => {
    before((done) => {
      answers.host = 'excel';
      answers.isManifestOnly = true;      

      helpers.run(path.join(__dirname, '../app'))
        .withArguments([])
        .withOptions({ name: projectEscapedName })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let manifestFileName = projectEscapedName  + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png'
      ];

      assert.file(expected);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel"
	 */
  describe('command line: --name "Display Name" --host excel', () => {    
    before((done) => {
      helpers.run(path.join(__dirname, '../app'))
        .withArguments([])
        .withOptions({ name: projectEscapedName }, { host: 'excel' })
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let manifestFileName = projectEscapedName  + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png'
      ];

      assert.file(expected);
      done();
    });
  });

	/**
	 * Test addin when user pass in argument
	 * "my-office-add-in excel manifest-only"
	 */
  describe('command line: manifest-only --name "Display Name" --host excel', () => {
    before((done) => {
      helpers.run(path.join(__dirname, '../app'))
        .withArguments(['manifest-only'])
        .withOptions({ name: projectEscapedName }, { host: 'excel'})
        .withPrompts(answers)
        .on('end', done);
    });

    it('creates expected files', (done) => {
      let manifestFileName = projectEscapedName  + '-manifest.xml';

      let expected = [
        manifestFileName,
        'package.json',
        'assets/icon-16.png',
        'assets/icon-32.png',
        'assets/icon-80.png',
        'assets/logo-filled.png'
      ];

      assert.file(expected);
      done();
    });
  });
});
