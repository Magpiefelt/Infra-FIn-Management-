// Excel Upload and Parsing Component
import React, { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog.jsx'
import { Alert, AlertDescription } from '@/components/ui/alert.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table.jsx'
import { Upload, FileSpreadsheet, CheckCircle, AlertCircle, X } from 'lucide-react'

// PFMT Data Extractor Component
export function PFMTDataExtractor({ project, onDataExtracted, onClose }) {
  const [file, setFile] = useState(null)
  const [extractedData, setExtractedData] = useState(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState(null)
  const [validationResults, setValidationResults] = useState(null)

  // Handle file drop
  const handleDrop = useCallback((e) => {
    e.preventDefault()
    const droppedFile = e.dataTransfer.files[0]
    if (droppedFile) {
      handleFileSelect(droppedFile)
    }
  }, [])

  const handleDragOver = useCallback((e) => {
    e.preventDefault()
  }, [])

  // Handle file selection
  const handleFileSelect = (selectedFile) => {
    if (!selectedFile) return

    // Validate file type
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel.sheet.macroEnabled.12'
    ]
    
    if (!validTypes.includes(selectedFile.type) && !selectedFile.name.endsWith('.xlsx') && !selectedFile.name.endsWith('.xlsm')) {
      setError('Please select a valid Excel file (.xlsx or .xlsm)')
      return
    }

    // Check file size (50MB limit)
    if (selectedFile.size > 50 * 1024 * 1024) {
      setError('File size must be less than 50MB')
      return
    }

    setFile(selectedFile)
    setError(null)
    processExcelFile(selectedFile)
  }

  // Process Excel file and extract PFMT data
  const processExcelFile = async (file) => {
    setIsProcessing(true)
    setError(null)

    try {
      const arrayBuffer = await file.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: 'array' })

      // Check if required sheets exist
      const requiredSheets = ['SP Fields']
      const availableSheets = workbook.SheetNames
      const missingSheets = requiredSheets.filter(sheet => !availableSheets.includes(sheet))

      if (missingSheets.length > 0) {
        throw new Error(`Missing required sheets: ${missingSheets.join(', ')}. Available sheets: ${availableSheets.join(', ')}`)
      }

      // Extract data from SP Fields sheet
      const spFieldsSheet = workbook.Sheets['SP Fields']
      
      // Extract key financial data from known cell locations
      const extractedData = {
        // Primary financial fields from SP Fields sheet
        taf: getCellValue(spFieldsSheet, 'B2'), // Total Approved Funding
        eac: getCellValue(spFieldsSheet, 'B6'), // Estimate at Completion
        currentYearCashflow: getCellValue(spFieldsSheet, 'B4'), // Current Year Cashflow Total
        currentYearTarget: getCellValue(spFieldsSheet, 'B7'), // Current Year Target
        
        // Additional fields for validation
        projectId: getCellValue(spFieldsSheet, 'B1'), // Project ID if available
        lastUpdated: new Date().toISOString(),
        
        // Metadata
        fileName: file.name,
        fileSize: file.size,
        extractedAt: new Date().toISOString(),
        availableSheets: availableSheets
      }

      // Calculate variances
      if (extractedData.taf && extractedData.eac) {
        extractedData.tafEacVariance = extractedData.taf - extractedData.eac
      }

      if (extractedData.currentYearTarget && extractedData.currentYearCashflow) {
        extractedData.cashflowVariance = extractedData.currentYearTarget - extractedData.currentYearCashflow
      }

      // Validate extracted data
      const validation = validateExtractedData(extractedData, project)
      setValidationResults(validation)
      setExtractedData(extractedData)

    } catch (err) {
      console.error('Error processing Excel file:', err)
      setError(`Error processing file: ${err.message}`)
    } finally {
      setIsProcessing(false)
    }
  }

  // Helper function to get cell value safely
  const getCellValue = (sheet, cellAddress) => {
    const cell = sheet[cellAddress]
    if (!cell) return null
    
    // Handle different cell types
    if (cell.t === 'n') return cell.v // Number
    if (cell.t === 's') return cell.v // String
    if (cell.t === 'b') return cell.v // Boolean
    if (cell.t === 'd') return cell.v // Date
    
    return cell.v
  }

  // Validate extracted data
  const validateExtractedData = (data, project) => {
    const issues = []
    const warnings = []

    // Check for required fields
    if (!data.taf || data.taf <= 0) {
      issues.push('Total Approved Funding (TAF) is missing or invalid')
    }
    if (!data.eac || data.eac <= 0) {
      issues.push('Estimate at Completion (EAC) is missing or invalid')
    }
    if (!data.currentYearCashflow || data.currentYearCashflow < 0) {
      issues.push('Current Year Cashflow is missing or invalid')
    }
    if (!data.currentYearTarget || data.currentYearTarget <= 0) {
      issues.push('Current Year Target is missing or invalid')
    }

    // Business rule validations
    if (data.taf && data.eac && Math.abs(data.tafEacVariance) > data.taf * 0.1) {
      warnings.push(`Large TAF vs EAC variance detected: ${formatCurrency(data.tafEacVariance)}`)
    }

    if (data.currentYearTarget && data.currentYearCashflow && Math.abs(data.cashflowVariance) > 1000000) {
      warnings.push(`Large cashflow variance detected: ${formatCurrency(data.cashflowVariance)} (>$1M)`)
    }

    // Check if EAC is reasonable compared to TAF
    if (data.taf && data.eac && data.eac > data.taf * 1.5) {
      warnings.push('EAC is significantly higher than TAF (>150%)')
    }

    return {
      isValid: issues.length === 0,
      issues,
      warnings,
      hasWarnings: warnings.length > 0
    }
  }

  // Format currency for display
  const formatCurrency = (value) => {
    if (value === null || value === undefined) return 'N/A'
    return new Intl.NumberFormat('en-CA', {
      style: 'currency',
      currency: 'CAD',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(value)
  }

  // Handle data import confirmation
  const handleImportData = () => {
    if (extractedData && validationResults?.isValid) {
      onDataExtracted(extractedData)
      onClose()
    }
  }

  return (
    <Dialog open={true} onOpenChange={onClose}>
      <DialogContent className="max-w-4xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="flex items-center space-x-2">
            <FileSpreadsheet className="h-5 w-5" />
            <span>Import PFMT Data</span>
          </DialogTitle>
          <DialogDescription>
            Upload your PFMT Excel workbook to automatically populate project financial data
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-6">
          {/* File Upload Area */}
          {!file && (
            <div
              className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center hover:border-gray-400 transition-colors"
              onDrop={handleDrop}
              onDragOver={handleDragOver}
            >
              <Upload className="mx-auto h-12 w-12 text-gray-400 mb-4" />
              <p className="text-lg font-medium text-gray-900 mb-2">
                Drop your PFMT Excel file here
              </p>
              <p className="text-sm text-gray-600 mb-4">
                or click to browse for files
              </p>
              <input
                type="file"
                accept=".xlsx,.xlsm"
                onChange={(e) => handleFileSelect(e.target.files[0])}
                className="hidden"
                id="file-upload"
              />
              <label htmlFor="file-upload">
                <Button variant="outline" className="cursor-pointer">
                  Select File
                </Button>
              </label>
              <p className="text-xs text-gray-500 mt-2">
                Supports .xlsx and .xlsm files up to 50MB
              </p>
            </div>
          )}

          {/* File Processing Status */}
          {file && (
            <Card>
              <CardHeader>
                <CardTitle className="text-sm">Selected File</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="flex items-center justify-between">
                  <div className="flex items-center space-x-2">
                    <FileSpreadsheet className="h-4 w-4" />
                    <span className="text-sm font-medium">{file.name}</span>
                    <Badge variant="outline">
                      {(file.size / 1024 / 1024).toFixed(1)} MB
                    </Badge>
                  </div>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => {
                      setFile(null)
                      setExtractedData(null)
                      setError(null)
                      setValidationResults(null)
                    }}
                  >
                    <X className="h-4 w-4" />
                  </Button>
                </div>
              </CardContent>
            </Card>
          )}

          {/* Processing Indicator */}
          {isProcessing && (
            <Alert>
              <AlertCircle className="h-4 w-4" />
              <AlertDescription>
                Processing Excel file and extracting PFMT data...
              </AlertDescription>
            </Alert>
          )}

          {/* Error Display */}
          {error && (
            <Alert variant="destructive">
              <AlertCircle className="h-4 w-4" />
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}

          {/* Extracted Data Preview */}
          {extractedData && (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <CheckCircle className="h-5 w-5 text-green-500" />
                  <span>Extracted PFMT Data</span>
                </CardTitle>
                <CardDescription>
                  Review the extracted data before importing
                </CardDescription>
              </CardHeader>
              <CardContent>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Field</TableHead>
                      <TableHead>Current Value</TableHead>
                      <TableHead>New Value</TableHead>
                      <TableHead>Change</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    <TableRow>
                      <TableCell className="font-medium">Total Approved Funding (TAF)</TableCell>
                      <TableCell>{formatCurrency(project.taf)}</TableCell>
                      <TableCell>{formatCurrency(extractedData.taf)}</TableCell>
                      <TableCell>
                        {extractedData.taf && project.taf ? 
                          formatCurrency(extractedData.taf - project.taf) : 'N/A'}
                      </TableCell>
                    </TableRow>
                    <TableRow>
                      <TableCell className="font-medium">Estimate at Completion (EAC)</TableCell>
                      <TableCell>{formatCurrency(project.eac)}</TableCell>
                      <TableCell>{formatCurrency(extractedData.eac)}</TableCell>
                      <TableCell>
                        {extractedData.eac && project.eac ? 
                          formatCurrency(extractedData.eac - project.eac) : 'N/A'}
                      </TableCell>
                    </TableRow>
                    <TableRow>
                      <TableCell className="font-medium">Current Year Cashflow</TableCell>
                      <TableCell>{formatCurrency(project.currentYearCashflow)}</TableCell>
                      <TableCell>{formatCurrency(extractedData.currentYearCashflow)}</TableCell>
                      <TableCell>
                        {extractedData.currentYearCashflow && project.currentYearCashflow ? 
                          formatCurrency(extractedData.currentYearCashflow - project.currentYearCashflow) : 'N/A'}
                      </TableCell>
                    </TableRow>
                    <TableRow>
                      <TableCell className="font-medium">Current Year Target</TableCell>
                      <TableCell>{formatCurrency(project.targetCashflow)}</TableCell>
                      <TableCell>{formatCurrency(extractedData.currentYearTarget)}</TableCell>
                      <TableCell>
                        {extractedData.currentYearTarget && project.targetCashflow ? 
                          formatCurrency(extractedData.currentYearTarget - project.targetCashflow) : 'N/A'}
                      </TableCell>
                    </TableRow>
                  </TableBody>
                </Table>

                {/* Calculated Variances */}
                <div className="mt-4 space-y-2">
                  <h4 className="font-medium">Calculated Variances</h4>
                  <div className="grid grid-cols-2 gap-4 text-sm">
                    <div>
                      <span className="text-gray-600">TAF vs EAC Variance:</span>
                      <span className={`ml-2 font-medium ${
                        extractedData.tafEacVariance > 0 ? 'text-green-600' : 'text-red-600'
                      }`}>
                        {formatCurrency(extractedData.tafEacVariance)}
                      </span>
                    </div>
                    <div>
                      <span className="text-gray-600">Cashflow Variance:</span>
                      <span className={`ml-2 font-medium ${
                        extractedData.cashflowVariance > 0 ? 'text-green-600' : 'text-red-600'
                      }`}>
                        {formatCurrency(extractedData.cashflowVariance)}
                      </span>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          )}

          {/* Validation Results */}
          {validationResults && (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  {validationResults.isValid ? (
                    <CheckCircle className="h-5 w-5 text-green-500" />
                  ) : (
                    <AlertCircle className="h-5 w-5 text-red-500" />
                  )}
                  <span>Validation Results</span>
                </CardTitle>
              </CardHeader>
              <CardContent>
                {validationResults.issues.length > 0 && (
                  <div className="mb-4">
                    <h4 className="font-medium text-red-600 mb-2">Issues Found:</h4>
                    <ul className="list-disc list-inside space-y-1 text-sm text-red-600">
                      {validationResults.issues.map((issue, index) => (
                        <li key={index}>{issue}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {validationResults.warnings.length > 0 && (
                  <div>
                    <h4 className="font-medium text-yellow-600 mb-2">Warnings:</h4>
                    <ul className="list-disc list-inside space-y-1 text-sm text-yellow-600">
                      {validationResults.warnings.map((warning, index) => (
                        <li key={index}>{warning}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {validationResults.isValid && validationResults.warnings.length === 0 && (
                  <p className="text-green-600 text-sm">
                    All validations passed. Data is ready for import.
                  </p>
                )}
              </CardContent>
            </Card>
          )}
        </div>

        <DialogFooter>
          <Button variant="outline" onClick={onClose}>
            Cancel
          </Button>
          <Button
            onClick={handleImportData}
            disabled={!extractedData || !validationResults?.isValid || isProcessing}
          >
            Import Data
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  )
}

