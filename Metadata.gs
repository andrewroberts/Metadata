// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 20171110
/* jshint asi: true */

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
var SCRIPT_VERSION = "v1.0"

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

var Metadata_ = (function myFunction(ns) {

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
    
  } // Subs_.load()

  /**
   * Add some meta data to a GSheet column
   *
   * @param {Sheet} sheet
   * @param {String} key The header name of the column (row 1)
   * @param {Number} startIndex 0-based index
   * 
   * @return {Object} result
   */
     
  ns.add = function(sheet, key, startIndex) {

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
    
    var spreadsheetId = sheet.getParent().getId()
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId)
    return result
  
  } // Metadata_.addMetaData()
    
  /**
   * Get some meta data from a GSheet
   *
   * @param {Sheet} sheet
   * @param {String} key 
   * 
   * @return {Object} result
   */

  ns.get = function(sheet, key) {
  
    ns.log.functionEntryPoint()
    ns.log.fine('key: ' + key)

    var callingfunction = 'Metadata_.get()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    Assert.assertString(key, callingfunction, 'key not a string')

    var spreadsheetId = sheet.getParent().getId()
    var meta
    
    try {
    
      meta = cSAM.SAM.searchByKey(spreadsheetId, key)
    
    } catch (error) {
    
      if (error.name === 'GoogleJsonResponseException') {
        
        ns.log.fine('The Sheets API cant be accessed')
        return []
      }   
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
   * Get a column's index 
   *
   * @param {Sheet} sheet
   * @param {String} key The header name of the column (row 1)
   * 
   * @return {Number} startIndex 0-based or -1
   */

  ns.getColumnIndex = function(sheet, key) {

    ns.log.functionEntryPoint()

    var tidied = ns.get(sheet, key)
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
   * @param {Sheet} sheet
   * @param {String} key 
   * 
   * @return {Object} result
   */
  
  ns.remove = function(sheet, key) {

    ns.log.functionEntryPoint()

    var callingfunction = 'Metadata_.remove()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
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
  
    var spreadsheetId = sheet.getParent().getId()
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, spreadsheetId);
    return result
    
  } // Metadata_.remove()  

  /**
   * Remove all the meta data from a GSheet
   *
   * @param {Sheet} sheet
   * @param {Object} columns - A list of column names
   */
  
  ns.removeAll = function(sheet, columns) {

    ns.log.functionEntryPoint()

    var callingfunction = 'Metadata_.removeAll()'
    Assert.assertObject(sheet, callingfunction, 'Bad "sheet" type')
    
    var spreadsheetId = sheet.getParent().getId()
    
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
  
  /**
   * Get the index (0-based) of this column
   *
   * First check if we can get it from column meta data, if not look in 
   * the header row for the column name
   *
   * @param {Object} 
   *   {Sheet} sheet
   *   {string} columnName
   *   {Array} headers [OPTIONAL, DEFAULT - got from row 1]
   *   {boolean} required [OPTIONAL, DEFAULT = true]
   *   {Boolean} useMeta [OPTIONAL, DEFAULT = true]
   *
   * @return {number} column index or -1
   */
    
  function getColumnIndex(config) {
      
    ns.log.functionEntryPoint()
    
    var sheet = config.sheet
    var columnName = config.columnName
    var headers = config.headers
    
    var required
    
    if (typeof config.required === 'undefined' || config.required) {
      required = true
    } else {
      required = false
    }
    
    var useMeta
    
    if (typeof config.useMeta === 'undefined' || config.useMeta) {
      useMeta = true
    } else {
      useMeta = false
    }
  
    ns.log.fine('columnName: %s', columnName)
    ns.log.fine('headers: %s', headers)
    ns.log.fine('required: %s', required)
    ns.log.fine('useMeta: %s', useMeta)
    
    var columnIndex = -1
    
    if (useMeta) {
    
      // First check if we can get it from column meta data
      
      columnIndex = MetaData_.getColumnIndex(sheet, columnName)
    }
    
    if (columnIndex === -1) {
      
      // Next try from the header
      
      if (typeof headers === 'undefined') {
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      }
      
      columnIndex = headers.indexOf(columnName)
      
      if (columnIndex === -1) {
      
        if (columnName === TASK_LIST_COLUMNS.TIMESTAMP) {
        
          // There may have been an old version where it was renamed to "Listed"
          columnIndex = headers.indexOf("Listed")
          
        } else if (columnName === TASK_LIST_COLUMNS.ID) {
        
          // Sometimes the 'ID' column header is accidentally deleted
          if (sheet.getRange('A1').getValue() === '') {        
            columnIndex = 0
          }
          
        } else if (columnName === TASK_LIST_COLUMNS.SUBJECT) {
        
          // These are used to be a regular error, so hard-coded (I'm that kinda guy!)
          
          columnIndex = headers.indexOf('Type of Service Request')
          
          if (columnIndex === -1) {
            columnIndex = headers.indexOf('Комментарий')
          }
        }
      }
      
      if (columnIndex !== -1) { 
      
        if (useMeta) {
      
          // Store meta data for this column in case it does get moved or renamed
          ns.add(sheet, columnName, columnIndex) 
        }
        
        ns.log.fine('columnIndex from headers: ' + columnIndex) 
      }
    }
    
    if (columnIndex === -1) {
      
      if (required) {
        
        Utils_.throwNoColumnError(columnName, headers)
        
      } else {
        
        ns.log.warning('No "' + columnName + '" column. Found: ' + JSON.stringify(headers))
      }
    }
    
    return columnIndex
    
  } // Metadata_.getColumnIndex()
  
  return ns
    
}) (Metadata_ || {})