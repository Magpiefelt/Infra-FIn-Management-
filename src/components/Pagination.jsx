// Pagination component for project dashboard
import React from 'react'

const Pagination = ({ 
  currentPage, 
  totalPages, 
  hasNext, 
  hasPrev, 
  onNextPage, 
  onPrevPage, 
  onPageChange,
  pageSize,
  onPageSizeChange,
  totalItems,
  className = ""
}) => {
  // Generate page numbers to display
  const getPageNumbers = () => {
    const pages = []
    const maxVisiblePages = 5
    
    if (totalPages <= maxVisiblePages) {
      // Show all pages if total is small
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i)
      }
    } else {
      // Show pages around current page
      let start = Math.max(1, currentPage - 2)
      let end = Math.min(totalPages, currentPage + 2)
      
      // Adjust if we're near the beginning or end
      if (currentPage <= 3) {
        end = Math.min(totalPages, 5)
      } else if (currentPage >= totalPages - 2) {
        start = Math.max(1, totalPages - 4)
      }
      
      // Add first page and ellipsis if needed
      if (start > 1) {
        pages.push(1)
        if (start > 2) {
          pages.push('...')
        }
      }
      
      // Add visible pages
      for (let i = start; i <= end; i++) {
        pages.push(i)
      }
      
      // Add ellipsis and last page if needed
      if (end < totalPages) {
        if (end < totalPages - 1) {
          pages.push('...')
        }
        pages.push(totalPages)
      }
    }
    
    return pages
  }

  const pageNumbers = getPageNumbers()
  const startItem = (currentPage - 1) * pageSize + 1
  const endItem = Math.min(currentPage * pageSize, totalItems)

  return (
    <div className={`flex flex-col sm:flex-row items-center justify-between gap-4 ${className}`}>
      {/* Items info */}
      <div className="text-sm text-gray-600">
        Showing {startItem} to {endItem} of {totalItems} projects
      </div>
      
      {/* Page size selector */}
      <div className="flex items-center gap-2">
        <label htmlFor="pageSize" className="text-sm text-gray-600">
          Show:
        </label>
        <select
          id="pageSize"
          value={pageSize}
          onChange={(e) => onPageSizeChange(parseInt(e.target.value))}
          className="border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        >
          <option value={5}>5</option>
          <option value={10}>10</option>
          <option value={20}>20</option>
          <option value={50}>50</option>
        </select>
        <span className="text-sm text-gray-600">per page</span>
      </div>

      {/* Pagination controls */}
      {totalPages > 1 && (
        <div className="flex items-center gap-1">
          {/* Previous button */}
          <button
            onClick={onPrevPage}
            disabled={!hasPrev}
            className={`px-3 py-1 rounded text-sm font-medium transition-colors ${
              hasPrev
                ? 'bg-white border border-gray-300 text-gray-700 hover:bg-gray-50'
                : 'bg-gray-100 border border-gray-200 text-gray-400 cursor-not-allowed'
            }`}
          >
            Previous
          </button>

          {/* Page numbers */}
          <div className="flex items-center gap-1">
            {pageNumbers.map((page, index) => (
              <React.Fragment key={index}>
                {page === '...' ? (
                  <span className="px-2 py-1 text-gray-400">...</span>
                ) : (
                  <button
                    onClick={() => onPageChange(page)}
                    className={`px-3 py-1 rounded text-sm font-medium transition-colors ${
                      page === currentPage
                        ? 'bg-blue-600 text-white'
                        : 'bg-white border border-gray-300 text-gray-700 hover:bg-gray-50'
                    }`}
                  >
                    {page}
                  </button>
                )}
              </React.Fragment>
            ))}
          </div>

          {/* Next button */}
          <button
            onClick={onNextPage}
            disabled={!hasNext}
            className={`px-3 py-1 rounded text-sm font-medium transition-colors ${
              hasNext
                ? 'bg-white border border-gray-300 text-gray-700 hover:bg-gray-50'
                : 'bg-gray-100 border border-gray-200 text-gray-400 cursor-not-allowed'
            }`}
          >
            Next
          </button>
        </div>
      )}
    </div>
  )
}

export default Pagination

