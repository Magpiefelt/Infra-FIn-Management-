// Report generation components
import React, { useState } from 'react'
import { Button } from '@/components/ui/button.jsx'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card.jsx'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog.jsx'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select.jsx'
import { Label } from '@/components/ui/label.jsx'
import { Badge } from '@/components/ui/badge.jsx'
import { Alert, AlertDescription } from '@/components/ui/alert.jsx'
import { FileText, Download, Calendar, TrendingUp, BarChart3, PieChart, FileSpreadsheet, Printer } from 'lucide-react'
import { formatDate, formatCurrency } from '../../utils/index.js'
import { useNotifications } from '../../hooks/index.js'
import { generateProjectReport } from '../../services/mockData.js'

export function ReportGenerator({ project }) {
  const [showReportDialog, setShowReportDialog] = useState(false)
  const [reportType, setReportType] = useState('monthly')
  const [generating, setGenerating] = useState(false)
  const [generatedReport, setGeneratedReport] = useState(null)
  const { showSuccess, showError } = useNotifications()

  const handleGenerateReport = async () => {
    setGenerating(true)
    try {
      const report = await generateProjectReport(project.id, reportType)
      setGeneratedReport(report)
      showSuccess('Report generated successfully!')
    } catch (error) {
      showError('Failed to generate report: ' + error.message)
    } finally {
      setGenerating(false)
    }
  }

  const handleDownloadReport = () => {
    if (!generatedReport) return

    // Create a comprehensive report content
    const reportContent = `
PROJECT FINANCIAL MANAGEMENT REPORT
Generated: ${formatDate(generatedReport.generatedDate)}
Report Type: ${generatedReport.reportType.toUpperCase()}

PROJECT SUMMARY
===============
Project Name: ${generatedReport.data.projectSummary.name}
Contractor: ${generatedReport.data.projectSummary.contractor}
Phase: ${generatedReport.data.projectSummary.phase}
Region: ${generatedReport.data.projectSummary.region}
Project Manager: ${generatedReport.data.projectSummary.projectManager}
Start Date: ${formatDate(generatedReport.data.projectSummary.startDate)}

FINANCIAL SUMMARY
================
Total Budget (TAF): ${formatCurrency(generatedReport.data.financialSummary.totalBudget)}
Amount Spent: ${formatCurrency(generatedReport.data.financialSummary.amountSpent)}
Estimate at Completion (EAC): ${formatCurrency(generatedReport.data.financialSummary.eac)}
Budget Variance: ${formatCurrency(generatedReport.data.financialSummary.variance)}
Current Year Cashflow: ${formatCurrency(generatedReport.data.financialSummary.currentYearCashflow)}
Target Cashflow: ${formatCurrency(generatedReport.data.financialSummary.targetCashflow)}

STATUS SUMMARY
==============
Schedule Status: ${generatedReport.data.statusSummary.scheduleStatus}
Budget Status: ${generatedReport.data.statusSummary.budgetStatus}
Report Status: ${generatedReport.data.statusSummary.reportStatus}
${generatedReport.data.statusSummary.scheduleReasonCode ? `Schedule Reason: ${generatedReport.data.statusSummary.scheduleReasonCode}` : ''}
${generatedReport.data.statusSummary.budgetReasonCode ? `Budget Reason: ${generatedReport.data.statusSummary.budgetReasonCode}` : ''}

MILESTONES
==========
${generatedReport.data.milestones.map(milestone => `
${milestone.name} - ${milestone.status.toUpperCase()}
  Planned: ${formatDate(milestone.plannedDate)}
  ${milestone.actualDate ? `Actual: ${formatDate(milestone.actualDate)}` : 'Actual: Pending'}
  ${milestone.notes ? `Notes: ${milestone.notes}` : ''}
`).join('')}

MONTHLY COMMENTS
===============
${generatedReport.data.monthlyComments || 'No comments provided'}

PREVIOUS HIGHLIGHTS
==================
${generatedReport.data.previousHighlights || 'No highlights provided'}

NEXT STEPS
==========
${generatedReport.data.nextSteps || 'No next steps provided'}

APPROVAL HISTORY
===============
${generatedReport.data.approvalHistory.submittedBy ? `Submitted by: ${generatedReport.data.approvalHistory.submittedBy} on ${formatDate(generatedReport.data.approvalHistory.submittedDate)}` : 'Not yet submitted'}
${generatedReport.data.approvalHistory.approvedBy ? `Approved by: ${generatedReport.data.approvalHistory.approvedBy} on ${formatDate(generatedReport.data.approvalHistory.approvedDate)}` : 'Not yet approved'}

MONTHLY ARCHIVE HISTORY
======================
${generatedReport.data.monthlyArchive.length > 0 ? 
  generatedReport.data.monthlyArchive.map(archive => `
${archive.reportingPeriod} - Archived on ${formatDate(archive.archivedDate)}
  Schedule Status: ${archive.data.scheduleStatus}
  Budget Status: ${archive.data.budgetStatus}
  TAF: ${formatCurrency(archive.data.taf)}
  EAC: ${formatCurrency(archive.data.eac)}
`).join('') : 'No archived data available'}

---
Report generated by PFMT Replacement System
${new Date().toISOString()}
    `.trim()

    // Create and download the file
    const blob = new Blob([reportContent], { type: 'text/plain' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${project.name.replace(/[^a-zA-Z0-9]/g, '_')}_${reportType}_report_${new Date().toISOString().slice(0, 10)}.txt`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)

    showSuccess('Report downloaded successfully!')
  }

  const reportTypes = [
    { value: 'monthly', label: 'Monthly Report', description: 'Standard monthly project status report' },
    { value: 'financial', label: 'Financial Summary', description: 'Detailed financial analysis and budget tracking' },
    { value: 'milestone', label: 'Milestone Report', description: 'Project milestone status and timeline analysis' },
    { value: 'comprehensive', label: 'Comprehensive Report', description: 'Complete project overview with all data' }
  ]

  return (
    <>
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center space-x-2">
            <FileText className="h-5 w-5" />
            <span>Report Generation</span>
          </CardTitle>
          <CardDescription>
            Generate and download project reports and analytics
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="p-4 border rounded-lg">
              <div className="flex items-center space-x-2 mb-2">
                <BarChart3 className="h-4 w-4 text-blue-600" />
                <span className="font-medium">Financial Status</span>
              </div>
              <div className="text-sm text-gray-600">
                <p>Budget: {formatCurrency(project.totalBudget)}</p>
                <p>Spent: {formatCurrency(project.amountSpent)}</p>
                <p>EAC: {formatCurrency(project.eac)}</p>
              </div>
            </div>
            
            <div className="p-4 border rounded-lg">
              <div className="flex items-center space-x-2 mb-2">
                <TrendingUp className="h-4 w-4 text-green-600" />
                <span className="font-medium">Project Status</span>
              </div>
              <div className="space-y-1">
                <Badge className={project.scheduleStatus === 'Green' ? 'bg-green-100 text-green-800' : 
                                project.scheduleStatus === 'Yellow' ? 'bg-yellow-100 text-yellow-800' : 
                                'bg-red-100 text-red-800'}>
                  Schedule: {project.scheduleStatus}
                </Badge>
                <Badge className={project.budgetStatus === 'Green' ? 'bg-green-100 text-green-800' : 
                                project.budgetStatus === 'Yellow' ? 'bg-yellow-100 text-yellow-800' : 
                                'bg-red-100 text-red-800'}>
                  Budget: {project.budgetStatus}
                </Badge>
              </div>
            </div>
          </div>

          <Dialog open={showReportDialog} onOpenChange={setShowReportDialog}>
            <DialogTrigger asChild>
              <Button className="w-full">
                <FileText className="h-4 w-4 mr-2" />
                Generate Report
              </Button>
            </DialogTrigger>
            <DialogContent>
              <DialogHeader>
                <DialogTitle>Generate Project Report</DialogTitle>
                <DialogDescription>
                  Select the type of report you want to generate for {project.name}
                </DialogDescription>
              </DialogHeader>
              
              <div className="space-y-4">
                <div>
                  <Label htmlFor="report-type">Report Type</Label>
                  <Select value={reportType} onValueChange={setReportType}>
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {reportTypes.map((type) => (
                        <SelectItem key={type.value} value={type.value}>
                          <div>
                            <div className="font-medium">{type.label}</div>
                            <div className="text-sm text-gray-600">{type.description}</div>
                          </div>
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                {generatedReport && (
                  <Alert>
                    <FileText className="h-4 w-4" />
                    <AlertDescription>
                      Report generated successfully! You can now download it or generate a new one.
                    </AlertDescription>
                  </Alert>
                )}
              </div>

              <DialogFooter>
                <Button variant="outline" onClick={() => setShowReportDialog(false)}>
                  Cancel
                </Button>
                {generatedReport && (
                  <Button variant="outline" onClick={handleDownloadReport}>
                    <Download className="h-4 w-4 mr-2" />
                    Download
                  </Button>
                )}
                <Button onClick={handleGenerateReport} disabled={generating}>
                  {generating ? 'Generating...' : 'Generate Report'}
                </Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>

          {project.monthlyArchive && project.monthlyArchive.length > 0 && (
            <div className="mt-4">
              <h4 className="font-medium mb-2">Available Historical Reports</h4>
              <div className="space-y-2 max-h-32 overflow-y-auto">
                {project.monthlyArchive.map((archive) => (
                  <div key={archive.id} className="flex items-center justify-between p-2 bg-gray-50 rounded text-sm">
                    <span>{archive.reportingPeriod} Report</span>
                    <div className="flex items-center space-x-2">
                      <span className="text-gray-600">{formatDate(archive.archivedDate)}</span>
                      <Button size="sm" variant="outline">
                        <Download className="h-3 w-3" />
                      </Button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </CardContent>
      </Card>
    </>
  )
}

export function SystemReporting({ projects }) {
  const [reportType, setReportType] = useState('summary')
  const [generating, setGenerating] = useState(false)
  const { showSuccess, showError } = useNotifications()

  const handleGenerateSystemReport = async () => {
    setGenerating(true)
    try {
      await new Promise(resolve => setTimeout(resolve, 1500)) // Simulate generation time
      
      const systemReportContent = `
PFMT SYSTEM SUMMARY REPORT
Generated: ${formatDate(new Date().toISOString())}

SYSTEM OVERVIEW
===============
Total Projects: ${projects.length}
Active Projects: ${projects.filter(p => p.status === 'Active').length}
Projects Requiring Updates: ${projects.filter(p => p.reportStatus === 'Update Required').length}
Approved Projects: ${projects.filter(p => p.directorApproved).length}

FINANCIAL OVERVIEW
==================
Total System Budget: ${formatCurrency(projects.reduce((sum, p) => sum + (p.totalBudget || 0), 0))}
Total Amount Spent: ${formatCurrency(projects.reduce((sum, p) => sum + (p.amountSpent || 0), 0))}
Total EAC: ${formatCurrency(projects.reduce((sum, p) => sum + (p.eac || 0), 0))}

PROJECT STATUS BREAKDOWN
========================
${projects.map(project => `
${project.name}
  Status: ${project.status}
  Report Status: ${project.reportStatus}
  Schedule: ${project.scheduleStatus}
  Budget: ${project.budgetStatus}
  Budget: ${formatCurrency(project.totalBudget)}
  Spent: ${formatCurrency(project.amountSpent)}
  EAC: ${formatCurrency(project.eac)}
  PM: ${project.projectManager}
`).join('')}

REGIONAL BREAKDOWN
==================
${Object.entries(projects.reduce((acc, p) => {
  acc[p.region] = (acc[p.region] || 0) + 1
  return acc
}, {})).map(([region, count]) => `${region}: ${count} projects`).join('\n')}

---
Report generated by PFMT Replacement System
${new Date().toISOString()}
      `.trim()

      // Download the system report
      const blob = new Blob([systemReportContent], { type: 'text/plain' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `PFMT_System_Report_${new Date().toISOString().slice(0, 10)}.txt`
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)

      showSuccess('System report generated and downloaded!')
    } catch (error) {
      showError('Failed to generate system report: ' + error.message)
    } finally {
      setGenerating(false)
    }
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center space-x-2">
          <PieChart className="h-5 w-5" />
          <span>System Reporting</span>
        </CardTitle>
        <CardDescription>
          Generate system-wide reports and analytics
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="text-center p-4 bg-blue-50 rounded-lg">
            <div className="text-2xl font-bold text-blue-600">{projects.length}</div>
            <div className="text-sm text-blue-800">Total Projects</div>
          </div>
          <div className="text-center p-4 bg-green-50 rounded-lg">
            <div className="text-2xl font-bold text-green-600">
              {projects.filter(p => p.status === 'Active').length}
            </div>
            <div className="text-sm text-green-800">Active</div>
          </div>
          <div className="text-center p-4 bg-yellow-50 rounded-lg">
            <div className="text-2xl font-bold text-yellow-600">
              {projects.filter(p => p.reportStatus === 'Update Required').length}
            </div>
            <div className="text-sm text-yellow-800">Pending Updates</div>
          </div>
          <div className="text-center p-4 bg-purple-50 rounded-lg">
            <div className="text-2xl font-bold text-purple-600">
              {formatCurrency(projects.reduce((sum, p) => sum + (p.totalBudget || 0), 0))}
            </div>
            <div className="text-sm text-purple-800">Total Budget</div>
          </div>
        </div>

        <Button 
          onClick={handleGenerateSystemReport} 
          disabled={generating}
          className="w-full"
        >
          {generating ? (
            <>
              <FileSpreadsheet className="h-4 w-4 mr-2 animate-pulse" />
              Generating System Report...
            </>
          ) : (
            <>
              <FileSpreadsheet className="h-4 w-4 mr-2" />
              Generate System Report
            </>
          )}
        </Button>
      </CardContent>
    </Card>
  )
}

