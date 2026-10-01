var assert = require('assert');
var carbone = require('../lib/index');
var path  = require('path');
var fs = require('fs');
var os = require('os');
var helper = require('../lib/helper');
var converter = require('../lib/converter');
var params = require('../lib/params');
var childProcess = require('child_process');
var exec = childProcess.exec;
var http = require('http');
var yazl = require('yazl');
var tempPath = path.join(__dirname, 'temp');

var defaultOptions = {
  pipeNamePrefix : '_carbone',
  factories      : 1,
  startFactory   : false,
  attempts       : 2,
  tempPath       : tempPath
};

describe('Converter', function () {
  beforeEach(function () {
    helper.rmDirRecursive(tempPath);
    fs.mkdirSync(tempPath, '0755');
  });
  afterEach(function () {
    helper.rmDirRecursive(tempPath);
  });
  after(function () {
    carbone.reset();
  });

  describe('shouldTheFactoryBeRestarted', function () {
    it('should return not restart LO if factoryMemoryFileSize = 0 or factoryMemoryThreshold = 0', function () {
      var _params = {
        factoryMemoryFileSize  : 0,
        factoryMemoryThreshold : 0
      };
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1000, 100), false);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1   , 1000000), false);
      _params.factoryMemoryFileSize = 10000;
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1   , 100000000), false);
      _params.factoryMemoryFileSize = 10000;
      _params.factoryMemoryThreshold = 0;
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1   , 100000000), false);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 100000000, 100000000), false);
    });
    it('should restart LO if the threshold is reached', function () {
      var _params = {
        factoryMemoryFileSize  : 2,
        factoryMemoryThreshold : 20
      };
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1000, 99), false);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1000, 100), true);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 1000, 101), true);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 10  , 101), true);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 10  , 1  ), true);
      helper.assert(converter.shouldTheFactoryBeRestarted(_params, 10  , 0  ), false);
    });
  });

  describe('init', function () {
    afterEach(function (done) {
      converter.exit(function () {
        converter.init(defaultOptions, done);
      });
    });
    /* it.skip('should be able to convert in docx, doc... ');
    it.skip('should accept temp directory with whitespaces');*/
    it('should start one conversion factory and return a object which describes the factories (pid, ...)', function (done) {
      converter.init({factories : 1, startFactory : true, tempPath : tempPath}, function (factories) {
        var _nbFactories = 0;
        var _cachePathRegex = new RegExp(tempPath);
        _nbFactories = Object.keys(factories).length;
        helper.assert(_nbFactories, 1);
        helper.assert(factories[0].mode, 'pipe');
        helper.assert(/_carbone/g.test(factories[0].pipeName), true);
        helper.assert(_cachePathRegex.test(factories[0].userCachePath), true);
        helper.assert(/[0-9]+/g.test(factories[0].pid), true);
        helper.assert(factories[0].isReady, true);
        helper.assert(factories[0].isConverting, false);
        done();
      });
    });
    it('should start 3 conversion factory and take into account the options object.\
             the factories should have different pids, and different cache path', function (done) {
      var _customOptions = {
        pipeNamePrefix : '_carboneTest',
        factories      : 3,
        startFactory   : true,
        tempPath       : tempPath
      };
      converter.init(_customOptions, function (factories) {
        var _nbFactories = 0;
        var _cachePathRegex = new RegExp(tempPath);
        _nbFactories = Object.keys(factories).length;
        helper.assert(_nbFactories, 3);

        helper.assert(factories[0].mode, 'pipe');
        helper.assert(/_carboneTest/g.test(factories[0].pipeName), true);
        helper.assert(_cachePathRegex.test(factories[0].userCachePath), true);
        helper.assert(/[0-9]+/g.test(factories[0].pid), true);
        helper.assert(factories[0].isReady, true);

        helper.assert(factories[2].mode, 'pipe');
        helper.assert(/_carboneTest/g.test(factories[0].pipeName), true);
        helper.assert(_cachePathRegex.test(factories[2].userCachePath), true);
        helper.assert(/[0-9]+/g.test(factories[2].pid), true);
        helper.assert(factories[2].isReady, true);

        helper.assert((factories[2].pid           !== factories[0].pid), true);
        helper.assert((factories[2].userCachePath !== factories[0].userCachePath), true);

        done();
      });
    });
  });

  describe('convertFile', function () {
    afterEach(function (done) {
      converter.exit(function () {
        setTimeout(function () {
          converter.init(defaultOptions, done);
        }, 1000);
      });
    });
    it('should render a pdf and start an conversion factory automatically if no factories exist', function (done) {
      var _pdfExpectedPath = path.resolve('./test/datasets/test_odt_render_static.pdf');
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _outputPath = path.join(tempPath, 'output1.pdf');
      converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
        var _result   = fs.readFileSync(resultPath);
        var _expected = fs.readFileSync(_pdfExpectedPath);
        helper.assert(resultPath, _outputPath);
        assert.equal(_result.slice(0, 4).toString(), '%PDF');
        assert.equal(_result.slice(8, 70).toString(), _expected.slice(8, 70).toString());
        done();
      });
    });
    it('should render JPG images', function (done) {
      const _magicNumberJPG = 'ffd8ff';
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _outputPath = path.join(tempPath, 'output181.jpg');
      converter.convertFile(_filePath, 'writer_jpg_Export', '', _outputPath, function (err, bufferSmallImagePath) {
        var _bufferSmallImage = fs.readFileSync(bufferSmallImagePath);
        assert.strictEqual(err, null);
        helper.assert(_bufferSmallImage.slice(0, 3).toString('hex'), _magicNumberJPG);
        var _outputPathBig = path.join(tempPath, 'output186.jpg');
        done();
      });
    });

    it('should render JPG images', function (done) {
      const _magicNumberJPG = 'ffd8ff';
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _outputPathLow = path.join(tempPath, 'output206.jpg');
      converter.convertFile(_filePath, 'writer_jpg_Export', '', _outputPathLow, function (err, bufferLowQualityPath) {
        assert.strictEqual(err, null);
        var _bufferLowQuality = fs.readFileSync(bufferLowQualityPath);
        helper.assert(_bufferLowQuality.slice(0, 3).toString('hex'), _magicNumberJPG);
        var _outputPathMax = path.join(tempPath, 'output212.jpg');
        done();
      });
    });

    it('should render a PNG image', function (done) {
      var _magicNumberPNG = '89504e470d0a1a0a';
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _outputNotCompressed = path.join(tempPath, 'output229.png');
      converter.convertFile(_filePath, 'writer_png_Export', '', _outputNotCompressed, function (err, bufferNotCompressed) {
        assert.strictEqual(err, null);
        var _bufferNotCompressed = fs.readFileSync(bufferNotCompressed);
        helper.assert(_bufferNotCompressed.slice(0, 8).toString('hex'), _magicNumberPNG);
        var _outputCompressed = path.join(tempPath, 'output235.png');
        done();
      });
    });

    it('should restart automatically the conversion factory if it crashes', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _outputPath = path.join(tempPath, 'output2.pdf');
      converter.init({factories : 1, startFactory : true, tempPath : tempPath}, function (factories) {
        converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
          helper.assert(err+'', 'null');
          var _result = fs.readFileSync(resultPath);
          assert.equal(_result.slice(0, 4).toString(), '%PDF');
          // kill LibreOffice thread
          process.kill(factories['0'].pid);
          // try another conversion
          // with Node v4, process.kill seems to be very slow. So I add a "setTimeout" but I'm not very happy with it
          setTimeout(function () {
            _outputPath = path.join(tempPath, 'output3.pdf');
            converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
              helper.assert(err+'', 'null');
              _result = fs.readFileSync(resultPath);
              assert.equal(_result.slice(0, 4).toString(), '%PDF');
              done();
            });
          }, 200);
        });
      });
    });
    it('should be fast to render a pdf with four Factories', function (done) {
      converter.init({factories : 4, startFactory : true, tempPath : tempPath}, function () {
        var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
        var _nbExecuted = 200;
        var _results = [];
        var _waitedResponse = _nbExecuted;
        var _start = new Date();
        for (var i = 0; i < _nbExecuted; i++) {
          var _outputPath = path.join(tempPath, 'output__'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            _waitedResponse--;
            if (!err) {
              _results.push(resultPath);
            }
            if (_waitedResponse === 0) {
              theEnd();
            }
          });
        }
        function theEnd () {
          var _end = new Date();
          var _elapsed = (_end.getTime() - _start.getTime())/_nbExecuted; // time in milliseconds
          console.log('\n\n Conversion to PDF Time Elapsed : '+_elapsed + ' ms per pdf for '+_nbExecuted+' conversions (usally around 65ms) \n\n\n');
          for (var i = 0; i < _results.length; i++) {
            var _result = fs.readFileSync(_results[i]);
            assert.equal(_result.slice(0, 4).toString(), '%PDF');
          }
          assert.equal((_elapsed < (350 * helper.CPU_PERFORMANCE_FACTOR)), true);
          done();
        }
      });
    });
    it('should be extremely robust. It should not loose any jobs even if the office or python thread crash randomly', function (done) {
      converter.init({factories : 4, startFactory : true, tempPath : tempPath, attempts : 10}, function (factories) {
        var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
        var _nbExecuted = 100;
        var _crashModulo = 28;
        var _results = [];
        var _waitedResponse = _nbExecuted;
        var _start = new Date();
        for (var i = 0; i < _nbExecuted; i++) {
          var _outputPath = path.join(tempPath, 'output__'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            if (_waitedResponse % _crashModulo === 0) {
              var _factoryId = Math.floor((Math.random()*4)); // (0 -> 3)
              var _threadChoice = Math.random(); // (0.0 -> 1.0)
              var _factory = factories[_factoryId];
              if (_factory) {
                if (_threadChoice > 0.5 && _factory.officeThread) {
                  _factory.officeThread.kill();
                }
                else if (_factory.pythonThread) {
                  _factory.pythonThread.kill();
                }
              }
            }
            _waitedResponse--;
            if (!err) {
              _results.push(resultPath);
            }
            if (_waitedResponse === 0) {
              theEnd();
            }
          });
        }
        function theEnd () {
          var _end = new Date();
          var _elapsed = (_end.getTime() - _start.getTime())/_nbExecuted; // time in milliseconds
          console.log('\n\n Conversion to PDF Time Elapsed : '+_elapsed + ' ms per pdf for '+_nbExecuted+' conversions with '+(_nbExecuted/_crashModulo).toFixed(0)+' crashes\n\n\n');
          for (var i = 0; i < _results.length; i++) {
            var _result = fs.readFileSync(_results[i]);
            assert.equal(_result.slice(0, 4).toString(), '%PDF');
          }
          assert.equal((_elapsed < 400), true);
          done();
        }
      });
    });
    it('should not restart the conversion factory if the document is corrupted. It should return an error message', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_corrupted.odt');
      converter.init({factories : 1, startFactory : true, tempPath : tempPath}, function (factories) {
        var _officePID = factories['0'].pid;
        var _outputPath = path.join(tempPath, 'output241.pdf');
        converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err) {
          assert.equal(err+'', 'Error: Could not open document');
          assert.equal(factories['0'].pid, _officePID);
          done();
        });
      });
    });
    it('should not restart the conversion factory if the document can be opened, but cannot be converted', function (done) {
      var _filePath = path.resolve('./test/datasets/test_spreadsheet.ods');
      converter.init({factories : 1, startFactory : true, tempPath : tempPath}, function (factories) {
        var _officePID = factories['0'].pid;
        var _outputPath = path.join(tempPath, 'output241.pdf');
        converter.convertFile(_filePath, 'MS Word 97', '', _outputPath, function (err) {
          assert.equal(err+'', 'Error: Could not convert document');
          assert.equal(factories['0'].pid, _officePID);
          done();
        });
      });
    });
    it('should still restart the conversion factory if the document could not be opened more than 10 times', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_corrupted.odt');
      var _nbAttemptMax = 10;
      var _nbAttempt = _nbAttemptMax;
      converter.init({factories : 1, startFactory : true, tempPath : tempPath, attempts : 1}, function (factories) {
        var _officePID = factories['0'].pid;
        for (var i = 0; i < _nbAttemptMax; i++) {
          var _outputPath = path.join(tempPath, 'output__'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err) {
            _nbAttempt--;
            assert.equal(/Could/.test(err), true);
            assert.equal(factories['0'].pid, _officePID);
            // the 10th conversion will restart LibreOffice
            if (_nbAttempt === 0) {
              _outputPath = path.join(tempPath, 'output265.pdf');
              converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err) {
                assert.equal(/Could/.test(err), true);
                assert.notEqual(factories['0'].pid, _officePID);
                _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
                _outputPath = path.join(tempPath, 'output284.pdf');
                converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
                  assert.equal(err, null);
                  var _result = fs.readFileSync(resultPath);
                  assert.equal(_result.slice(0, 4).toString(), '%PDF');
                  done();
                });
              });
            }
          });
        }
      });
    });
    it('should not restart the conversion factory if there is at least one success between 10 fails\
      it should not crash if formatOption is undefined', function (done) {
      var _filePath = '';
      var _filePathKO = path.resolve('./test/datasets/test_odt_render_corrupted.odt');
      var _filePathOK = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _nbAttemptMax = 15;
      var _nbAttempt = _nbAttemptMax;
      converter.init({factories : 1, startFactory : true, tempPath : tempPath, attempts : 1}, function (factories) {
        var _officePID = factories['0'].pid;
        for (var i = 0; i < _nbAttemptMax; i++) {
          if (i===7) {
            _filePath = _filePathOK;
          }
          else {
            _filePath = _filePathKO;
          }
          var _outputPath = path.join(tempPath, 'output__'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', undefined, _outputPath, function () {
            _nbAttempt--;
            assert.equal(factories['0'].pid, _officePID);
            if (_nbAttempt === 0) {
              done();
            }
          });
        }
      });
    });
  });

  describe('LO Memory monitoring and report timeout', function () {
    const totatMemoryAvailableMB     = os.totalmem() / 1024 / 1024;
    const factoryMemoryThreshold     = 8; // 8 %
    const nbConversionBoundary       = 5; // below it os ok, above it restarts LO
    const nbConversionForRestart     = nbConversionBoundary + 1;
    const nbConversionWithoutRestart = nbConversionBoundary - 1;
    // adapt factoryMemoryFileSize according to the system memory to make the test machine-independant
    const factoryMemoryFileSize      = totatMemoryAvailableMB * factoryMemoryThreshold / (nbConversionBoundary * 100);
    const memoryOptions = {
      factories              : 1,
      startFactory           : true,
      tempPath               : tempPath,
      factoryMemoryFileSize  : factoryMemoryFileSize,
      factoryMemoryThreshold : factoryMemoryThreshold
    };

    afterEach(function (done) {
      converter.exit(done);
      carbone.reset();
    });

    it('should not reach the memory threshold and the factory process shouldn\'t be killed', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      converter.init(memoryOptions, function (factories) {
        const _officePID = factories['0'].pid;
        for (let i = 0; i <= nbConversionWithoutRestart; i++) {
          var _outputPath = path.join(tempPath, 'output_'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            helper.assert(err+'', 'null');
            var _result = fs.readFileSync(resultPath);
            assert.equal(_result.slice(0, 4).toString(), '%PDF');
            assert.equal(factories['0'].pid, _officePID);
            if (i === nbConversionWithoutRestart) {
              done();
            }
          });
        }
      });
    });

    it('should reach the memory threshold  and the factory process should be killed', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      converter.init(memoryOptions, function (factories) {
        const _officePID = factories['0'].pid;
        for (let i = 0; i <= nbConversionForRestart; i++) {
          var _outputPath = path.join(tempPath, 'output_'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            helper.assert(err+'', 'null');
            var _result = fs.readFileSync(resultPath);
            assert.equal(_result.slice(0, 4).toString(), '%PDF');
            if (i === nbConversionForRestart) {
              assert.notEqual(factories['0'].pid, _officePID);
              done();
            }
          });
        }
      });
    });

    it('should reach the memory threshold and 4 factories should be killed', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      memoryOptions.factories = 4;
      converter.init(memoryOptions, function (factories) {
        const _officePIDs = [factories['0'].pid, factories['1'].pid, factories['2'].pid, factories['3'].pid];
        const _maxLoops = nbConversionForRestart * memoryOptions.factories;
        for (let i = 0; i <= _maxLoops; i++) {
          var _outputPath = path.join(tempPath, 'output_'+i+'.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            helper.assert(err+'', 'null');
            var _result = fs.readFileSync(resultPath);
            assert.equal(_result.slice(0, 4).toString(), '%PDF');
            if (i === _maxLoops) {
              for (let i = 0; i < _officePIDs.length; i++) {
                assert.notEqual(factories[i+''].pid, _officePIDs[i]);
              }
              done();
            }
          });
        }
      });
    });

    it('should not timeout and should not kill the factory process. (timeout disabled)', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      memoryOptions.converterFactoryTimeout = 0;
      converter.init(memoryOptions, function (factories) {
        const _officePID = factories['0'].pid;
        var _outputPath = path.join(tempPath, 'output_413.pdf');
        converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
          helper.assert(err+'', 'null');
          var _result = fs.readFileSync(resultPath);
          assert.equal(_result.slice(0, 4).toString(), '%PDF');
          assert.equal(factories['0'].pid, _officePID);
          done();
        });
      });
    });

    it('should timeout if the report is too long to convert, return an error and should be able to convert again after', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      memoryOptions.converterFactoryTimeout = 5;
      converter.init(memoryOptions, function (factories) {
        const _officePID = factories['0'].pid;
        var _outputPath = path.join(tempPath, 'output_429.pdf');
        converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
          helper.assert(err instanceof Error, true);
          helper.assert(err+'', 'Error: Document conversion timeout reached ('+params.converterFactoryTimeout+' ms)');
          helper.assert(resultPath, _outputPath);
          params.converterFactoryTimeout = 10000;
          _outputPath = path.join(tempPath, 'output_435.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            const _newOfficePIDAfter = factories['0'].pid;
            helper.assert(_officePID !== _newOfficePIDAfter, true);
            helper.assert(err+'', 'null');
            var _result = fs.readFileSync(resultPath);
            assert.equal(_result.slice(0, 4).toString(), '%PDF');
            done();
          });
        });
      });
    });

    it('should convert one file, another with timeout, and another without timeout', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      memoryOptions.converterFactoryTimeout = 5;
      converter.init(memoryOptions, function (factories) {
        const _officePID = factories['0'].pid;
        params.converterFactoryTimeout = 10000;
        var _outputPath = path.join(tempPath, 'output_454.pdf');
        converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
          const _newOfficePID = factories['0'].pid;
          helper.assert(_officePID, _newOfficePID);
          helper.assert(err+'', 'null');
          var _result = fs.readFileSync(resultPath);
          assert.equal(_result.slice(0, 4).toString(), '%PDF');
          params.converterFactoryTimeout = 2;
          _outputPath = path.join(tempPath, 'output_4262.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
            helper.assert(err instanceof Error, true);
            helper.assert(err+'', 'Error: Document conversion timeout reached ('+params.converterFactoryTimeout+' ms)');
            helper.assert(resultPath, _outputPath);
            params.converterFactoryTimeout = 10000;
            _outputPath = path.join(tempPath, 'output_468.pdf');
            converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err, resultPath) {
              const _newOfficePIDAfter = factories['0'].pid;
              helper.assert(_officePID !== _newOfficePIDAfter, true);
              helper.assert(err+'', 'null');
              _result = fs.readFileSync(resultPath);
              assert.equal(_result.slice(0, 4).toString(), '%PDF');
              done();
            });
          });
        });
      });
    });
  });

  describe('exit', function () {
    before(function () {
      var _tempContent = helper.walkDirSync(tempPath);
      if (_tempContent.length > 0) {
        console.error(`Warning - ${_tempContent.length} files hasn't been deleted at the end of the previous tests.`);
        helper.rmDirRecursive(tempPath);
      }
    });

    afterEach(function (done) {
      converter.exit(function () {
        converter.init(defaultOptions, done);
      });
    });
    it('should start one conversion factory and clean the temp directory', function (done) {
      converter.init({factories : 1, startFactory : true, tempPath : tempPath}, function () {
        converter.exit(function () {
          setTimeout(function () {
            var _tempContent = helper.walkDirSync(tempPath);
            assert.equal(_tempContent.length, 0);
            done();
          }, 4000);
        });
      });
    });
    it('should not delete files which come from another carbone instance or another program', function (done) {
      var _otherFile = path.join(tempPath, 'ovni.sql');
      fs.writeFileSync(_otherFile, 'test');
      converter.init({factories : 1, startFactory : true, tempPath : tempPath}, function () {
        converter.exit(function () {
          setTimeout(function () {
            var _tempContent = helper.walkDirSync(tempPath);
            assert.equal(_tempContent.length, 1);
            assert.equal(_tempContent[0], _otherFile);
            exec('rm -rf '+_otherFile, done);
          }, 4000);
        });
      });
    });
  });

  describe('environment of LibreOffice and Python', function () {
    var _originalSpawn = childProcess.spawn;
    var _spawned = [];
    beforeEach(function (done) {
      _spawned = [];
      childProcess.spawn = function (command, args, options) {
        _spawned.push({ command : command, options : options });
        return _originalSpawn.apply(this, arguments);
      };
      converter.exit(done);
    });
    afterEach(function (done) {
      childProcess.spawn = _originalSpawn;
      delete process.env.CARBONE_TEST_SECRET;
      params.converterEnv = {};
      converter.exit(function () {
        converter.init(defaultOptions, done);
      });
    });
    it('should keep only the variables LibreOffice and Python need, then add converterEnv', function () {
      var _env = converter.buildEnvironment({
        PATH              : '/usr/bin',
        Path              : 'C:\\Windows',
        HOME              : '/home/carbone',
        LANG              : 'tr_TR.UTF-8',
        LC_ALL            : 'C.UTF-8',
        SAL_USE_VCLPLUGIN : 'svp',
        MONGODB_URI       : 'mongodb://user:password@host/db',
        JWT_SECRET        : 'secret',
        AWS_SECRET_ACCESS_KEY : 'secret'
      }, { FONTCONFIG_SYSROOT : '/fonts' });
      assert.deepStrictEqual(_env, {
        PATH               : '/usr/bin',
        Path               : 'C:\\Windows',
        HOME               : '/home/carbone',
        LANG               : 'tr_TR.UTF-8',
        LC_ALL             : 'C.UTF-8',
        SAL_USE_VCLPLUGIN  : 'svp',
        FONTCONFIG_SYSROOT : '/fonts'
      });
    });
    it('should give everything when converterEnv is process.env, as before', function () {
      process.env.CARBONE_TEST_SECRET = 'visible';
      var _env = converter.buildEnvironment(process.env, process.env);
      assert.strictEqual(_env.CARBONE_TEST_SECRET, 'visible');
    });
    it('should not give the secrets of the parent process to LibreOffice and Python', function (done) {
      process.env.CARBONE_TEST_SECRET = 'not-for-libreoffice';
      converter.init({ factories : 1, startFactory : true, tempPath : tempPath }, function (factories) {
        assert.strictEqual(_spawned.length, 2);
        _spawned.forEach(function (spawned) {
          assert.ok(spawned.options.env.PATH, spawned.command + ' gets PATH');
          assert.strictEqual(spawned.options.env.CARBONE_TEST_SECRET, undefined, spawned.command + ' does not get the secret');
        });
        // On Linux, read the environment of the running processes too (macOS does not show it for signed apps)
        if (process.platform === 'linux') {
          [factories[0].pid, factories[0].pythonThread.pid].forEach(function (pid) {
            var _environ = fs.readFileSync('/proc/' + pid + '/environ', 'utf8');
            assert.strictEqual(_environ.indexOf('CARBONE_TEST_SECRET'), -1);
          });
        }
        done();
      });
    });
  });

  describe('start errors', function () {
    var _originalSpawn = childProcess.spawn;
    afterEach(function (done) {
      childProcess.spawn = _originalSpawn;
      converter.exit(function () {
        // the failed start blocks new starts for 5 s: wait before the next tests start LibreOffice
        setTimeout(function () {
          converter.init(defaultOptions, done);
        }, 5500);
      });
    });
    it('should not crash nor loop when LibreOffice cannot be started', function (done) {
      this.timeout(20000);
      var _starts = 0;
      childProcess.spawn = function (command, args, options) {
        _starts++;
        return _originalSpawn.call(this, path.join(tempPath, 'missing-executable'), args, options);
      };
      converter.init({ factories : 1, startFactory : true, tempPath : tempPath });
      setTimeout(function () {
        // one office and one python process tried once, then nothing during the retry delay
        assert.strictEqual(_starts, 2);
        done();
      }, 2000);
    });
  });

  describe('links of a document', function () {
    var PICTURE = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=', 'base64');
    var _server = null;
    var _requests = 0;
    before(function (done) {
      _server = http.createServer(function (req, res) {
        _requests++;
        res.writeHead(200, { 'Content-Type' : 'image/png' });
        res.end(PICTURE);
      });
      _server.listen(0, '127.0.0.1', done);
    });
    after(function (done) {
      _server.close(done);
    });
    afterEach(function (done) {
      params.converterBlockExternalLinks = false;
      converter.exit(function () {
        converter.init(defaultOptions, done);
      });
    });

    /* A DOCX whose only picture is linked to an address (not stored in the document) */
    function writeLinkedPictureDocx (url, outputPath, callback) {
      var _zip = new yazl.ZipFile();
      _zip.addBuffer(Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'), '[Content_Types].xml');
      _zip.addBuffer(Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'), '_rels/.rels');
      _zip.addBuffer(Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + url + '" TargetMode="External"/></Relationships>'), 'word/_rels/document.xml.rels');
      _zip.addBuffer(Buffer.from('<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:link="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>'), 'word/document.xml');
      _zip.outputStream.pipe(fs.createWriteStream(outputPath)).on('close', callback);
      _zip.end();
    }

    function convertLinkedPicture (blockLinks, callback) {
      converter.exit(function () {
        converter.init({ factories : 1, tempPath : tempPath, converterBlockExternalLinks : blockLinks }, function () {
          var _input = path.join(tempPath, 'linked.docx');
          writeLinkedPictureDocx('http://127.0.0.1:' + _server.address().port + '/picture.png', _input, function () {
            _requests = 0;
            converter.convertFile(_input, 'writer_pdf_Export', '', path.join(tempPath, 'linked.pdf'), function (err) {
              setTimeout(function () {
                callback(err, _requests);
              }, 300);
            });
          });
        });
      });
    }

    it('should load a linked picture by default', function (done) {
      convertLinkedPicture(false, function (err, requests) {
        assert.strictEqual(err, null);
        assert.ok(requests > 0, 'LibreOffice requests the picture');
        done();
      });
    });
    it('should not open the address of a linked picture if converterBlockExternalLinks is true', function (done) {
      convertLinkedPicture(true, function (err, requests) {
        assert.strictEqual(err, null);
        assert.strictEqual(requests, 0);
        done();
      });
    });
  });

  describe('factoryRecycleAfter', function () {
    afterEach(function (done) {
      converter.exit(function () {
        converter.init(defaultOptions, done);
      });
      carbone.reset();
    });
    it('should restart LibreOffice after factoryRecycleAfter documents, not before', function (done) {
      var _filePath = path.resolve('./test/datasets/test_odt_render_static.odt');
      var _options = { factories : 1, startFactory : true, tempPath : tempPath, factoryRecycleAfter : 3, factoryMemoryThreshold : 0 };
      converter.init(_options, function (factories) {
        var _firstPID = factories['0'].pid;
        var _pids = [];
        var _convert = function (index) {
          if (index === 4) {
            // documents 1 to 3 in the first process, the 4th in a new one
            assert.deepStrictEqual(_pids.slice(0, 3), [_firstPID, _firstPID, _firstPID]);
            assert.notEqual(_pids[3], _firstPID);
            return done();
          }
          var _outputPath = path.join(tempPath, 'recycle_' + index + '.pdf');
          converter.convertFile(_filePath, 'writer_pdf_Export', '', _outputPath, function (err) {
            helper.assert(err+'', 'null');
            _pids.push(factories['0'].pid);
            // the process is killed right after the 3rd document: wait for the new one before the next document
            setTimeout(function () {
              _convert(index + 1);
            }, index === 2 ? 1000 : 0);
          });
        };
        _convert(0);
      });
    });
  });

  describe('profiles of stopped processes', function () {
    afterEach(function (done) {
      converter.exit(function () {
        converter.init(defaultOptions, done);
      });
    });
    it('should remove the LibreOffice profiles of stopped Node processes and keep the others', function (done) {
      var _profile = function (name) {
        var _dir = path.join(tempPath, name);
        fs.mkdirSync(path.join(_dir, 'user'), { recursive : true });
        fs.writeFileSync(path.join(_dir, 'user', 'registrymodifications.xcu'), '<oor:items/>');
        return _dir;
      };
      // pid 999999 is above the default pid limit of Linux and macOS: no process has it
      var _stopped = _profile('_office_' + params.uidPrefix + '_1700000000000_999999_0');
      var _running = _profile('_office_' + params.uidPrefix + '_1700000000000_' + process.ppid + '_0');
      var _notCarbone = _profile('_office_backup');
      converter.init({ factories : 1, startFactory : true, tempPath : tempPath }, function () {
        assert.strictEqual(fs.existsSync(_stopped), false);
        assert.strictEqual(fs.existsSync(_running), true);
        assert.strictEqual(fs.existsSync(_notCarbone), true);
        done();
      });
    });
  });

});
