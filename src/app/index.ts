/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as appInsights from 'applicationinsights';
import * as chalk from 'chalk';
import * as _ from 'lodash';
import * as opn from 'opn';
import * as uuid from 'uuid/v4';
import * as yosay from 'yosay';
import * as yo from 'yeoman-generator';
let insight = appInsights.getClient('1ced6a2f-b3b2-4da5-a1b8-746512fbc840');

// Remove unwanted tags
delete insight.context.tags['ai.cloud.roleInstance'];
delete insight.context.tags['ai.device.osVersion'];
delete insight.context.tags['ai.device.osArchitecture'];
delete insight.context.tags['ai.device.osPlatform'];

module.exports = yo.extend({
  /**
   * Setup the generator
   */
  constructor: function () {
    yo.apply(this, arguments);

    const currentDir = path.resolve('.');
    const dirName = path.basename(currentDir);

    this.argument('framework', { type: String, required: false });

    this.option('output', {
      type: String,
      required: false,
      desc: 'Location to place the generated output. If not specified, uses the current directory.'
    });

    this.option('name', {
      type: String,
      required: false,
      desc: 'Name of the add-in. If not specified, uses the directory name of the output location.'
    });

    this.option('host', {
      type: String,
      required: false,      
      desc: 'Office app which will host the add-in.'
    });

    this.option('skip-install', {
      type: Boolean,
      required: false,
      desc: 'Skip running `npm install` post scaffolding.'
    });

    this.option('js', {
      type: Boolean,
      required: false,
      desc: 'Use JavaScript templates instead of TypeScript.'
    });
  },

  /**
   * Generator initalization
   */
  initializing: function () {
    let message = `Welcome to the ${chalk.bold.green('Office Add-in')} generator, by ${chalk.bold.green('@OfficeDev')}! Let\'s create a project together!`;
    this.log(yosay(message));
    this.project = {};
  },

  /**
   * Prompt users for options
   */
  prompting: async function () {
    try {
      let jsTemplates = getDirectories(this.templatePath('js'));
      let tsTemplates = getDirectories(this.templatePath('ts'));
      const hosts = getDirectories(this.templatePath('hosts'));
      const wantPrompt: boolean = (this.options.template === null); // prompt only if no arguments are specified
  
      // if prompts are not desired, provide defaults for options
      if (!wantPrompt) {
        // if output location is not specified, use the current folder
        if (!this.options.output) {
          this.options.output = path.resolve('.');
        }

        // if name is not specified, use the folder name of the output location
        if (!this.options.name) {
          this.options.name = path.basename(this.options.output);
        }
      }


      /** begin prompting */
      /** whether to create a new folder for the project */
      const startForFolder = getTime();
      const answerForFolder = await this.prompt([{
        name: 'folder',
        message: 'Would you like to create a new subfolder for your project?',
        type: 'confirm',
        default: false,
        when: this.options.output == null
      }]);
      const durationForFolder = getTimeSpan(startForFolder);

      /** name for the project */
      let startForName = getTime();
      let answerForName = await this.prompt([{
        name: 'name',
        type: 'input',
        message: 'What do you want to name your add-in?',
        default: 'My Office Add-in',
        when: this.options.name == null
      }]);
      let durationForName = getTimeSpan(startForName);

      /** office client application that can host the addin */
      let startForHost = getTime();
      let answerForHost = await this.prompt([{
        name: 'host',
        message: 'Which Office client application would you like to support?',
        type: 'list',
        default: 'Excel',
        choices: hosts.map(host => ({ name: host, value: host })),
        when: wantPrompt && (this.options.host == null)
      }]);
      let durationForHost = getTimeSpan(startForHost);

      /** set flag for manifest-only to prompt accordingly later */
      let startForManifestOnly = getTime();
      let answerForManifestOnly = await this.prompt([{
        name: 'isManifestOnly',
        message: 'Would you like to create a new add-in?',
        type: 'list',
        default: false,
        choices: [
          {
            name: 'Yes, I need to create a new web app and manifest file for my add-in.',
            value: false
          },
          {
            name: 'No, I already have a web app and only need a manifest file for my add-in.',
            value: true
          }
        ],
        when: wantPrompt && (this.options.framework == null)
      }]);
      let durationForManifestOnly = getTimeSpan(startForManifestOnly);

      /**
       * Configure user input to have correct values
       */
      this.project = {
        folder: answerForFolder.folder,
        output: this.options.output || null,
        name: this.options.name || answerForName.name,
        host: this.options.host || answerForHost.host || 'excel',        
        framework: this.options.framework || null,
        isManifestOnly: answerForManifestOnly.isManifestOnly
      };

      if (answerForManifestOnly.isManifestOnly) {
        this.project.framework = 'manifest-only';
      }

      if (this.options.framework != null) {
        if (this.options.framework === 'manifest-only') {
          this.project.isManifestOnly = true;
        } else {
          this.project.isManifestOnly = false;
        }
      }

      /** askForTs and askForFramework will only be triggered if it's not a manifest-only project */
      /** use TypeScript for the project */
      let startForTs = getTime();
      let answerForTs = await this.prompt([{
          name: 'ts',
          type: 'confirm',
          message: 'Would you like to use TypeScript?',
          default: true,
          when: (this.options.js == null) && (!this.project.isManifestOnly) && (this.options.framework !== 'react')
      }]);
      let durationForTs = getTimeSpan(startForTs);

      if (!(this.options.js == null)) {
        this.project.ts = !this.options.js;
      }
      else {
        this.project.ts = answerForTs.ts || false;
      }

      if (this.options.framework === 'react') {
        this.project.ts = true;
      }

      /** technology used to create the addin (html / angular / etc) */
      let startForFramework = getTime();
      let answerForFramework = await this.prompt([
        {
          name: 'framework',
          message: 'Choose a framework:',
          type: 'list',
          default: 'react',
          choices: tsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
          when: (this.project.framework == null) && this.project.ts && !this.options.js && !answerForManifestOnly.isManifestOnly
        },
        {
          name: 'framework',
          message: 'Choose a framework:',
          type: 'list',
          default: 'jquery',
          choices: jsTemplates.map(template => ({ name: _.capitalize(template), value: template })),
          when: (this.project.framework == null) && !this.project.ts && this.options.js && !answerForManifestOnly.isManifestOnly
        }
      ]);
      let durationForFramework = getTimeSpan(startForFramework);

      if (!(this.options.framework == null)) {
        this.project.framework = this.options.framework;
      }
      else if (this.project.isManifestOnly === true) {
        this.project.framework = 'manifest-only';
      }
      else {
        this.project.framework = answerForFramework.framework;
      }

      let startForResourcePage = getTime();
      this.log('\nFor more information and resources on your next steps, we have created a resource.html file in your project.');
      let answerForOpenResourcePage = await this.prompt([
        {
          name: 'open',
          type: 'confirm',
          message: 'Would you like to open it now while we finish creating your project?',
          default: true,
          when: wantPrompt
        }
      ]);
      let endForResourcePage = getTime();
      let durationForResourcePage = getTimeSpan(startForResourcePage, endForResourcePage);
      this.project.isResourcePageOpened = answerForOpenResourcePage.open;
      this.project.duration = getTimeSpan(startForFolder, endForResourcePage);

      /** appInsights logging */
      if (this.project.folder) {
        insight.trackEvent('Folder', { CreatedSubFolder: this.project.folder.toString() }, { durationForFolder });
      }
      insight.trackEvent('Name', { Name: this.project.name }, { durationForName });
      insight.trackEvent('Host', { Host: this.project.host }, { durationForHost });
      insight.trackEvent('IsManifestOnly', { IsManifestOnly: this.project.isManifestOnly.toString() }, { durationForManifestOnly });
      insight.trackEvent('IsResourcePageOpened', { IsResourcePageOpened: this.project.isResourcePageOpened.toString() }, { durationForResourcePage });

      if (this.project.isManifestOnly === false) {
        insight.trackEvent('IsTs', { IsTs: this.project.ts.toString() }, { durationForTs });
        insight.trackEvent('Framework', { Framework: this.project.framework }, { durationForFramework });
      }
    } catch (err) {
      insight.trackException(new Error('Prompting Error: ' + err));
    }

  },

  /**
   * save configs & config project
   */
  configuring: function () {
    try {
      this.project.projectInternalName = _.kebabCase(this.project.name);
      this.project.projectDisplayName = this.project.name;
      this.project.projectId = uuid();
      this.project.hostInternalName = _.toLower(this.project.host);

      if (this.project.output) {
        this.destinationRoot(this.project.output);
      } else if (this.project.folder) {
        this.destinationRoot(this.project.projectInternalName);
      }

      let duration = this.project.duration;
      insight.trackEvent('App_Data', { AppID: this.project.projectId, Host: this.project.host, Framework: this.project.framework, isTypeScript: this.project.ts.toString() }, { duration });
    } catch (err) {
      insight.trackException(new Error('Configuration Error: ' + err));
    }
  },

  writing: {
    copyFiles: function () {
      try {
        let language = this.project.ts ? 'ts' : 'js';

        /** Show type of project creating in progress */
        if (this.project.framework !== 'manifest-only') {
          this.log('\n----------------------------------------------------------------------------------\n');
          this.log(`      Creating ${chalk.bold.green(this.project.projectDisplayName)} add-in using ${chalk.bold.magenta(language)} and ${chalk.bold.cyan(this.project.framework)}\n`);
          this.log('----------------------------------------------------------------------------------\n\n');
        }
        else {
          this.log('----------------------------------------------------------------------------------\n');
          this.log(`      Creating manifest for ${chalk.bold.green(this.project.projectDisplayName)} add-in\n`);
          this.log('----------------------------------------------------------------------------------\n\n');
        }

        /** Copy the manifest */
        this.fs.copyTpl(this.templatePath(`hosts/${this.project.host}/manifest.xml`), this.destinationPath(`${this.project.projectInternalName}-manifest.xml`), this.project);

        if (this.project.framework === 'manifest-only') {
          this.fs.copyTpl(this.templatePath(`manifest-only/**`), this.destinationPath(), this.project);
        }
        else {
          /** Copy the base template */
          this.fs.copy(this.templatePath(`${language}/base/**`), this.destinationPath());

          /** Copy the framework specific overrides */
          this.fs.copyTpl(this.templatePath(`${language}/${this.project.framework}/**`), this.destinationPath(), this.project);
        }
      } catch (err) {
        insight.trackException(new Error('File Copy Error: ' + err));
      }
    }
  },

  install: function () {
    try {
      if (this.project.isResourcePageOpened) {
        opn(`resource.html`);
      }
      if (this.options['skip-install']) {
        this.installDependencies({
          npm: false,
          bower: false,
          callback: this._postInstallHints.bind(this)
        });
      }
      else {
        this.installDependencies({
          npm: true,
          bower: false,
          callback: this._exitProcess.bind(this)
        });
      }
    } catch (err) {
      insight.trackException(new Error('Installation Error: ' + err));
      process.exitCode = 1;
    }
  },

  _postInstallHints: function () {
    /** Next steps and npm commands */
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this.log(`      ${chalk.green('Congratulations!')} Your add-in has been created! Your next steps:\n`);
    this.log(`      1. Launch your local web server via ${chalk.inverse(' npm start ')} (you may also need to`);
    this.log(`         trust the Self-Signed Certificate for the site if you haven't done that)`);
    this.log(`      2. Sideload the add-in into your Office application.\n`);
    this.log(`      Please refer to resource.html in your project for more information.`);
    this.log(`      Or visit our repo at: https://github.com/officeDev/generator-office \n`);
    this.log('----------------------------------------------------------------------------------------------------------\n');
    this._exitProcess();
  },

  _exitProcess: function () {
    process.exit();
  }
} as any);

function getDirectories(root) {
  return fs.readdirSync(root).filter(file => {
    if (file === 'base') {
      return false;
    }
    return fs.statSync(path.join(root, file)).isDirectory();
  });
}

function getFiles(root) {
  return fs.readdirSync(root).filter(file => {
    return !(fs.statSync(path.join(root, file)).isDirectory());
  });
}

function getTime() {
  return (new Date()).getTime();
}

function getTimeSpan(startTime, endTime = getTime()) {
  return (endTime - startTime) / 1000;
}

function updateHostNames(arr, key, newval) {
  let match = _.some(arr, _.method('match', key));
  if (match) {
    let index = _.indexOf(arr, key);
    arr.splice(index, 1, newval);
  }
}
