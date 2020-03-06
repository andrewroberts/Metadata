// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 20171110
/* jshint asi: true */
// Script version: 1.1 Local build

(function() {"use strict"})()

// MetaData
// ========
//
// Library for managing GSheet metadata. Mainly builds on cSAM (goo.gl/sa6VGp) 
// to provide an easier way to manage metadata on GSheet columns
//
// http://ramblings.mcpher.com/Home/excelquirks/totallyunscripted/samlib
//
// Copyright 2018 Andrew Roberts
//  
//  Licensed under the Apache License, Version 2.0 (the "License");
//  you may not use this file except in compliance with the License.
//  You may obtain a copy of the License at
//  
//  //    http://www.apache.org/licenses/LICENSE-2.0
//  
//  Unless required by applicable law or agreed to in writing, software
//  distributed under the License is distributed on an "AS IS" BASIS,
//  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//  See the License for the specific language governing permissions and
//  limitations under the License.

// TODO
// ----
//
// - sort out column index
// - pass sheet on load

// Global Config
// -------------

var SCRIPT_NAME    = "Metadata"
var SCRIPT_VERSION = "v1.2"

// Private Config
// --------------

// Dummy logging object
var Log_ = {
  functionEntryPoint : function() {},
  finest             : function() {},
  finer              : function() {},
  fine               : function() {},
  info               : function() {},
  warning            : function() {},
}

// User Guide
// ----------

/*

See the script bound to the Metadata Test sheet (https://goo.gl/F1dZcJ) for 
example on using Metadata.

  .
  .
  .
  // Setup: Get a Metadata object - See Metadata_ description below
  var metadata = Metadata.get({
    log         : BBLog.getLog(),                                     
  })
  .
  .
  .
  // 
  var response = metadata(...)
  
*/

// Public Code
// -----------

/**
 * The MetadataGetConfig event parameter object
 *
 * @typedef {object} MetadataGetConfig
 * @property {BBLog} log - A logging service with the same API as BBLog (github.com/andrewroberts/BBLog) [OPTIONAL]
 */

/**
 * Public method to get a subscription object
 *  
 * @param {MetadataGetConfig} config - {@link MetadataGetConfig} 
 *
 * @return {MetaData_} object 
 */
 
function load(config) {
  return Metadata_.load(config)
}

// Private Code
// ------------

var Metadata_ = (function(ns) {

  ns.log         = null

  /**
   * Private method to get a Metadata object
   *
   * @param {MetadataGetConfig} config - {@link MetadataGetConfig} 
   *
   * @return {MetaData_} object 
   */
   
  ns.load = function(config) {

    if (typeof config !== 'undefined') {
      ns.log = config.log || Log_
    } else {
      ns.log = Log_
    }
    
    ns.log.functionEntryPoint()    
    return Object.create(this)
    
  } // Metadata_.load()

  /**
   * Add some meta data to a GSheet column
   *
   * @param {Sheet} sheet
   * @param {String} key The header name of the column (row 1)
   * @param {Number} startIndex 0-based index
   * @param {String} spreadsheetId (optional, default - parent of sheet)
   * 
   * @return {Object} result
   */
     
  ns.add = function(sheet, key, startIndex, spreadsheetId) {

    ns.log.functionEntryPoint()
    
    ns.log.fine('key: ' + key)
    ns.log.fine('startIndex: ' + startIndex)

    var callingfunction = 'Metadata_.add()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    Assert.assertString(key, callingfunction, 'key not a string')
    Assert.assertNumber(startIndex, callingfunction, 'startIndex not a Number')
  
    var requests = [{
         
      // CreateDeveloperMetadataRequest
      createDeveloperMetadata: {
      
        // DeveloperMetaData
        developerMetadata: {
        
          // DeveloperMetaDataLocation with column scope  
          metadataKey: key,
          metadataValue: JSON.stringify({
            writtenBy:Session.getActiveUser().getEmail(),
            createdAt:new Date().getTime()
          }),
          location: {  
            dimensionRange: {
              sheetId:sheet.getSheetId(),
              dimension:"COLUMNS",
              startIndex:startIndex,    
              endIndex:startIndex + 1   
            }
          },
          visibility:"DOCUMENT"      
        }
      }
    }]
    
    if (spreadsheetId === undefined) {
      spreadsheetId = sheet.getParent().getId()
    }
    
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId)
    return result
  
  } // Metadata_.add()
    
  /**
   * Get some meta data from a GSheet
   *
   * @param {Sheet} sheet
   * @param {String} key 
   * @param {String} spreadsheetId (optional, default - parent of sheet)   
   * 
   * @return {Object} result or null
   */

  ns.get = function(sheet, key, spreadsheetId) {
  
    ns.log.functionEntryPoint()
    ns.log.fine('key: ' + key)

    var callingfunction = 'Metadata_.get()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    Assert.assertString(key, callingfunction, 'key not a string')

    if (spreadsheetId === undefined) {
      spreadsheetId = sheet.getParent().getId()
    }
    
    var meta = cSAM.SAM.searchByKey(spreadsheetId, key)
    
    if (meta === null) {
      return null
    }
    
    var tidied = cSAM.SAM.tidyMatched(meta)
    
    if (tidied.length === 0) {
      ns.log.fine('No meta data')
    } else {
      ns.log.fine('Got meta data')
    }

    return tidied
    
  } // Metadata_.get()

  /**
   * Get all the column metadata for a spreadsheet
   *
   * @param {String} spreadsheetId (optional, default - active)   
   * 
   * @return {Object} tidied
   */

  ns.getAllColumns = function(spreadsheetId) {

    ns.log.functionEntryPoint()
    ns.log.fine('spreadsheetId: ' + spreadsheetId)
    
    if (spreadsheetId === undefined) {
      spreadsheetId = SpreadsheetApp.getActive().getId()
    }

    var meta = Sheets.Spreadsheets.DeveloperMetadata.search({
      dataFilters: {
        developerMetadataLookup: {
          locationType: 'COLUMN'
        }
      }}, 
      spreadsheetId)
      
    var tidied = cSAM.SAM.tidyMatched(meta) 
    
    return tidied
      
  } // Metadata_.getAllColumns()

  /**
   * Get a column's start index 
   *
   * @param {Sheet} sheet
   * @param {String} key The header name of the column (row 1)
   * @param {String} spreadsheetId (optional, default - parent of sheet)         
   * 
   * @return {Number} startIndex 0-based or -1
   */

  ns.getColumnIndex = function(sheet, key, spreadsheetId) {

    ns.log.functionEntryPoint()

    var tidied = ns.get(sheet, key, spreadsheetId)
    
    if (tidied === null) {
      return -1
    }
    
    var startIndex = -1

    if (tidied.length !== 0) {

      if (tidied.length > 1) {
        ns.log.warning('Using the first column meta data (of ' + tidied.length + ') for "' + key + '"')
      }
      
      startIndex = tidied[0].location.dimensionRange.startIndex 
      ns.log.fine('startIndex: ' + startIndex)
      
    } else {
    
      ns.log.fine('No meta data for start index')
    } 
      
    return startIndex

  } // Metadata_.getColumnIndex()

  /**
   * Remove some meta data from a GSheet
   *
   * @param {Sheet} sheet (optional if spreadsheetId defined)
   * @param {String} key 
   * @param {String} spreadsheetId (optional, default - parent of sheet)      
   * 
   * @return {Object} result
   */
  
  ns.remove = function(sheet, key, spreadsheetId) {

    ns.log.functionEntryPoint()

    var callingfunction = 'Metadata_.remove()'
    Assert.assertString(key, callingfunction, 'key not a string')

    // Get all the things and delete them in one go
    var requests = [{
      deleteDeveloperMetadata: {
        dataFilter:{
          developerMetadataLookup: {
            metadataKey: key
          }
        }
      }
    }]
  
    if (spreadsheetId === undefined) {
      spreadsheetId = sheet.getParent().getId()
    }
    
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId);
    return result
    
  } // Metadata_.remove()  

  /**
   * Remove all the meta data from a GSheet
   *
   * @param {Sheet} sheet (optional if spreadsheetId defined)
   * @param {Object} columns - A list of column names
   * @param {String} spreadsheetId (optional, default - parent of sheet)   
   */
  
  ns.removeAll = function(sheet, columns, spreadsheetId) {

    ns.log.functionEntryPoint()

    var callingfunction = 'Metadata_.removeAll()'
    Assert.assertObject(columns, callingfunction, 'Bad "columns" type')
    
    if (spreadsheetId === undefined) {
      spreadsheetId = sheet.getParent().getId()
    }
        
    for (var key in columns) {

      if (columns.hasOwnProperty(key)) {

        // Get all the things and delete them in one go
        var requests = [{
          deleteDeveloperMetadata: {
            dataFilter:{
              developerMetadataLookup: {
                metadataKey: columns[key]
              }
            }
          }
        }]
    
        var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId)
        ns.log.fine('result: ' + result)    
      }
    }
    
  } // Metadata_.removeAll()  
    
  return ns
    
}) (Metadata_ || {})