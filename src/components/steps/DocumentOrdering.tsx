import { useState, useMemo, useEffect } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  TouchSensor,
  useSensor,
  useSensors,
} from '@dnd-kit/core'
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from '@dnd-kit/sortable'
import { CSS } from '@dnd-kit/utilities'
import {
  ArrowRightIcon,
  ArrowLeftIcon,
  Bars3Icon,
  TrashIcon,
  ChevronUpIcon,
  ChevronDownIcon,
  DocumentCheckIcon
} from '@heroicons/react/24/outline'
import { cn, generateId } from '@/utils'
import { documentTypeConfig } from '@/data/documents'
import { documentService } from '@/services/documentService'
import type { SelectedDocument, DocumentType, Document } from '@/types'

interface DocumentOrderingProps {
  formData: Partial<import('@/types').ProjectFormData>
  selectedDocuments: SelectedDocument[]
  onUpdateSelectedDocuments: (documents: SelectedDocument[]) => void
  onNext: () => void
  onPrevious: () => void
}

interface SortableItemProps {
  document: SelectedDocument
  index: number
  totalCount: number
  onRemove: (id: string) => void
  onMoveUp: (id: string) => void
  onMoveDown: (id: string) => void
}

function SortableItem({ document, index, totalCount, onRemove, onMoveUp, onMoveDown }: SortableItemProps) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: document.id })

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  }

  return (
    <motion.div
      ref={setNodeRef}
      style={style}
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -20 }}
      transition={{ delay: index * 0.1 }}
      className={cn(
        "bg-white dark:bg-gray-800 rounded-xl border-2 p-6 transition-all duration-200",
        isDragging
          ? "border-primary-500 shadow-2xl scale-105 rotate-2 z-50 opacity-90 bg-opacity-90"
          : "border-gray-200 dark:border-gray-700 hover:border-primary-300 hover:shadow-md"
      )}
    >
      <div className="flex items-center gap-4">
        {/* Drag Handle */}
        <button
          {...attributes}
          {...listeners}
          className="cursor-grab active:cursor-grabbing p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded-lg transition-colors touch-none"
          aria-label="Drag to reorder"
        >
          <Bars3Icon className="w-5 h-5 text-gray-400" />
        </button>

        {/* Order Number */}
        <div className="flex items-center justify-center w-8 h-8 bg-primary-100 dark:bg-primary-900/30 text-primary-600 dark:text-primary-400 rounded-full font-semibold text-sm">
          {index + 1}
        </div>

        {/* Document Info */}
        <div className="flex-1 min-w-0">
          <h3 className="font-semibold text-gray-900 dark:text-white truncate">
            {document.document.name}
          </h3>
        </div>

        {/* Actions */}
        <div className="flex items-center gap-1">
          {/* Move Up Button */}
          <button
            onClick={() => onMoveUp(document.id)}
            disabled={index === 0}
            className={cn(
              "btn btn-ghost btn-sm p-2 transition-colors",
              index === 0
                ? "opacity-30 cursor-not-allowed"
                : "hover:bg-blue-50 dark:hover:bg-blue-900/20 text-blue-600 dark:text-blue-400"
            )}
            title="Move up"
          >
            <ChevronUpIcon className="w-4 h-4" />
          </button>

          {/* Move Down Button */}
          <button
            onClick={() => onMoveDown(document.id)}
            disabled={index === totalCount - 1}
            className={cn(
              "btn btn-ghost btn-sm p-2 transition-colors",
              index === totalCount - 1
                ? "opacity-30 cursor-not-allowed"
                : "hover:bg-blue-50 dark:hover:bg-blue-900/20 text-blue-600 dark:text-blue-400"
            )}
            title="Move down"
          >
            <ChevronDownIcon className="w-4 h-4" />
          </button>

          {/* Remove Button */}
          <button
            onClick={() => onRemove(document.id)}
            className="btn btn-ghost btn-sm p-2 text-red-500 hover:text-red-700 hover:bg-red-50 dark:hover:bg-red-900/20"
            title="Remove document"
          >
            <TrashIcon className="w-4 h-4" />
          </button>
        </div>
      </div>
    </motion.div>
  )
}

export default function DocumentOrdering({
  formData,
  selectedDocuments,
  onUpdateSelectedDocuments,
  onNext,
  onPrevious,
}: DocumentOrderingProps) {
  const [searchTerm, setSearchTerm] = useState('')
  const [filterType, setFilterType] = useState<DocumentType | 'all'>('all')
  const [availableDocuments, setAvailableDocuments] = useState<Document[]>([])
  const [loading, setLoading] = useState(false)

  useEffect(() => {
    const loadDocuments = async () => {
      try {
        setLoading(true)
        console.log('Loading documents for product type:', formData.productType)
        
        // First, try to get documents from the service
        const docs = await documentService.getAllDocuments()
        console.log('All documents from service:', docs)
        
        // Then filter by product type if specified
        const filteredDocs = formData.productType 
          ? docs.filter(doc => doc.productType === formData.productType)
          : docs
          
        console.log('Filtered documents:', filteredDocs)
        setAvailableDocuments(filteredDocs)
      } catch (error) {
        console.error('Error loading documents:', error)
      } finally {
        setLoading(false)
      }
    }
    
    loadDocuments()
    
    // Also set up a reload listener in case documents are added/updated
    const handleStorageChange = (e: StorageEvent) => {
      if (e.key === 'documents_updated') {
        loadDocuments()
      }
    }
    
    window.addEventListener('storage', handleStorageChange)
    return () => window.removeEventListener('storage', handleStorageChange)
  }, [formData.productType])

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 2,
      },
    }),
    useSensor(TouchSensor, {
      activationConstraint: {
        delay: 50,
        tolerance: 5,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  )

  // Filter and search documents for selection
  const filteredDocuments = useMemo(() => {
    console.log('Filtering documents. Current documents:', availableDocuments)
    console.log('Current filters - search:', searchTerm, 'type:', filterType)
    
    return availableDocuments.filter(doc => {
      // Log each document being checked
      console.log('Checking document:', {
        name: doc.name,
        type: doc.type,
        productType: doc.productType,
        matchesSearch: doc.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
                      (doc.description || '').toLowerCase().includes(searchTerm.toLowerCase()),
        matchesFilter: filterType === 'all' || doc.type === filterType
      })
      
      const matchesSearch = searchTerm === '' || 
                         doc.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         (doc.description || '').toLowerCase().includes(searchTerm.toLowerCase())
                         
      const matchesFilter = filterType === 'all' || doc.type === filterType
      
      return matchesSearch && matchesFilter
    })
  }, [availableDocuments, searchTerm, filterType, formData.productType])

  // Get unique document types for filter
  const documentTypes = useMemo(() => {
    const types = Array.from(new Set(availableDocuments.map(doc => doc.type)))
    return types.sort((a, b) => {
      const configA = documentTypeConfig[a as DocumentType]
      const configB = documentTypeConfig[b as DocumentType]
      const priorityA = configA?.priority ?? 99
      const priorityB = configB?.priority ?? 99
      return priorityA - priorityB
    })
  }, [availableDocuments])

  // Filter only selected documents and sort by order
  const sortedDocuments = selectedDocuments
    .filter(doc => doc.selected)
    .sort((a, b) => a.order - b.order)

  // Check if document is selected
  const isDocumentSelected = (documentId: string): boolean => {
    return selectedDocuments.some(doc => doc.document.id === documentId && doc.selected)
  }

  // Toggle document selection
  const toggleDocument = (document: typeof availableDocuments[0]) => {
    const isSelected = isDocumentSelected(document.id)

    if (isSelected) {
      const updated = selectedDocuments.map(doc =>
        doc.document.id === document.id ? { ...doc, selected: false } : doc
      )
      onUpdateSelectedDocuments(updated)
    } else {
      const existingDoc = selectedDocuments.find(doc => doc.document.id === document.id)
      if (existingDoc) {
        const updated = selectedDocuments.map(doc =>
          doc.document.id === document.id ? { ...doc, selected: true } : doc
        )
        onUpdateSelectedDocuments(updated)
      } else {
        const newSelectedDoc: SelectedDocument = {
          id: generateId(),
          document,
          order: selectedDocuments.length,
          selected: true,
        }
        onUpdateSelectedDocuments([...selectedDocuments, newSelectedDoc])
      }
    }
  }

  function handleDragEnd(event: any) {
    const { active, over } = event

    if (active.id !== over?.id) {
      const oldIndex = sortedDocuments.findIndex(doc => doc.id === active.id)
      const newIndex = sortedDocuments.findIndex(doc => doc.id === over.id)

      const reorderedDocs = arrayMove(sortedDocuments, oldIndex, newIndex)

      // Update order property
      const updatedDocs = reorderedDocs.map((doc, index) => ({
        ...doc,
        order: index
      }))

      const unselectedDocs = selectedDocuments.filter(doc => !doc.selected)
      onUpdateSelectedDocuments([...updatedDocs, ...unselectedDocs])
    }
  }

  const removeDocument = (documentId: string) => {
    const updatedDocuments = selectedDocuments.map(doc =>
      doc.id === documentId ? { ...doc, selected: false } : doc
    )
    onUpdateSelectedDocuments(updatedDocuments)
  }

  // Move document up in order
  const moveDocumentUp = (documentId: string) => {
    const currentIndex = sortedDocuments.findIndex(doc => doc.id === documentId)
    if (currentIndex > 0) {
      const newDocuments = [...sortedDocuments]
      const [movedDoc] = newDocuments.splice(currentIndex, 1)
      newDocuments.splice(currentIndex - 1, 0, movedDoc)

      // Update orders
      const updatedDocuments = selectedDocuments.map(doc => {
        const newIndex = newDocuments.findIndex(newDoc => newDoc.id === doc.id)
        if (newIndex !== -1) {
          return { ...doc, order: newIndex + 1 }
        }
        return doc
      })
      onUpdateSelectedDocuments(updatedDocuments)
    }
  }

  // Move document down in order
  const moveDocumentDown = (documentId: string) => {
    const currentIndex = sortedDocuments.findIndex(doc => doc.id === documentId)
    if (currentIndex < sortedDocuments.length - 1) {
      const newDocuments = [...sortedDocuments]
      const [movedDoc] = newDocuments.splice(currentIndex, 1)
      newDocuments.splice(currentIndex + 1, 0, movedDoc)

      // Update orders
      const updatedDocuments = selectedDocuments.map(doc => {
        const newIndex = newDocuments.findIndex(newDoc => newDoc.id === doc.id)
        if (newIndex !== -1) {
          return { ...doc, order: newIndex + 1 }
        }
        return doc
      })
      onUpdateSelectedDocuments(updatedDocuments)
    }
  }

  const canProceed = sortedDocuments.length > 0

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.6 }}
      className="max-w-6xl mx-auto"
    >
      <div className="card p-8 lg:p-12">
        {/* Header */}
        <div className="text-center mb-10">
          <motion.div
            initial={{ scale: 0 }}
            animate={{ scale: 1 }}
            transition={{ delay: 0.2, type: "spring", stiffness: 200 }}
            className="w-16 h-16 bg-gradient-accent rounded-2xl flex items-center justify-center mx-auto mb-6 shadow-lg"
          >
            <DocumentCheckIcon className="w-8 h-8 text-white" />
          </motion.div>

          <h2 className="text-3xl font-bold text-gray-900 dark:text-white mb-4">
            Select & Arrange Documents
          </h2>
          <p className="text-lg text-gray-600 dark:text-gray-300 max-w-2xl mx-auto mb-6">
            First, select the documents you want to include. Then arrange their order below.
          </p>
        </div>

        {/* Document Selection Section */}
        <div className="bg-gray-50 dark:bg-gray-800 p-6 rounded-lg border border-gray-200 dark:border-gray-700 mb-8">
          <h3 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
            Select Documents <span className="text-red-500">*</span>
          </h3>
          <p className="text-sm text-gray-600 dark:text-gray-400 mb-6">
            Choose which documents to include in your PDF packet.
          </p>


          {/* Selected Count */}
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="mb-6 p-4 bg-white dark:bg-gray-700 rounded-lg border border-gray-200 dark:border-gray-600"
          >
            <p className="text-sm font-medium text-gray-700 dark:text-gray-300">
              {sortedDocuments.length} document{sortedDocuments.length !== 1 ? 's' : ''} selected
              {sortedDocuments.length > 0 && (
                <span className="ml-2 text-cyan-600 dark:text-cyan-400">
                  ✓ Ready to arrange
                </span>
              )}
            </p>
          </motion.div>

          {/* Loading State */}
          {loading && (
            <div className="text-center py-8">
              <p className="text-gray-500 dark:text-gray-400">Loading documents...</p>
            </div>
          )}

          {/* Documents Grid */}
          {!loading && (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 max-h-96 overflow-y-auto">
            <AnimatePresence>
              {filteredDocuments.map((document, index) => {
                const isSelected = isDocumentSelected(document.id)

                return (
                  <motion.div
                    key={document.id}
                    layout
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.9 }}
                    transition={{ delay: index * 0.05 }}
                    className={cn(
                      'relative cursor-pointer transition-all duration-200',
                      'border-2 rounded-lg p-4 bg-white dark:bg-gray-700',
                      isSelected
                        ? 'border-primary-500 bg-primary-50 dark:bg-primary-900/20 shadow-md'
                        : 'border-gray-200 dark:border-gray-600 hover:border-primary-300 hover:shadow-sm'
                    )}
                    onClick={() => toggleDocument(document)}
                  >
                    {/* Selection Checkbox */}
                    <div className="absolute top-3 right-3">
                      <div
                        className={cn(
                          "w-6 h-6 rounded border-2 flex items-center justify-center transition-all",
                          isSelected
                            ? 'bg-primary-500 border-primary-500'
                            : 'bg-white dark:bg-gray-600 border-gray-300 dark:border-gray-500'
                        )}
                      >
                        {isSelected && (
                          <motion.span
                            initial={{ scale: 0 }}
                            animate={{ scale: 1 }}
                            transition={{ type: 'spring', stiffness: 500, damping: 30 }}
                            className="text-white text-sm font-bold"
                          >
                            ✓
                          </motion.span>
                        )}
                      </div>
                    </div>

                    {/* Document Info */}
                    <h4 className="font-medium text-sm text-gray-900 dark:text-white">
                      {document.name}
                    </h4>
                  </motion.div>
                )
              })}
            </AnimatePresence>
            </div>
          )}

          {/* No Results */}
          {filteredDocuments.length === 0 && (
            <div className="text-center py-8">
              <p className="text-gray-500 dark:text-gray-400">
                {formData.productType
                  ? `No documents available for ${formData.productType === 'structural-floor' ? 'Structural Floor Panel' : 'Underlayment'}.`
                  : 'No documents found.'}
              </p>
            </div>
          )}
        </div>

        {/* Document Ordering Section */}
        {sortedDocuments.length > 0 && (
          <div className="bg-gray-50 dark:bg-gray-800 p-6 rounded-lg border border-gray-200 dark:border-gray-700 mb-8">
            <h3 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
              Arrange Document Order
            </h3>

            {/* Instructions */}
            <div className="bg-blue-50 dark:bg-blue-900/20 rounded-lg p-4 mb-6">
              <h4 className="font-semibold text-blue-900 dark:text-blue-100 mb-2 text-sm">How to reorder documents:</h4>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs text-blue-800 dark:text-blue-200">
                <div className="flex items-center gap-2">
                  <Bars3Icon className="w-4 h-4" />
                  <span>Drag handle (⋮⋮⋮) to reorder</span>
                </div>
                <div className="flex items-center gap-2">
                  <ChevronUpIcon className="w-3 h-3" />
                  <ChevronDownIcon className="w-3 h-3" />
                  <span>Arrow buttons to move one position</span>
                </div>
                <div className="flex items-center gap-2">
                  <TrashIcon className="w-4 h-4" />
                  <span>Trash icon to remove</span>
                </div>
              </div>
              <p className="text-xs text-blue-700 dark:text-blue-300 mt-2 italic">
                Your cover page will always appear first.
              </p>
            </div>

            {/* Document List */}
            <DndContext
              sensors={sensors}
              collisionDetection={closestCenter}
              onDragEnd={handleDragEnd}
            >
              <SortableContext
                items={sortedDocuments.map(doc => doc.id)}
                strategy={verticalListSortingStrategy}
              >
                <div className="space-y-4">
                  <AnimatePresence>
                    {sortedDocuments.map((document, index) => (
                      <SortableItem
                        key={document.id}
                        document={document}
                        index={index}
                        totalCount={sortedDocuments.length}
                        onRemove={removeDocument}
                        onMoveUp={moveDocumentUp}
                        onMoveDown={moveDocumentDown}
                      />
                    ))}
                  </AnimatePresence>
                </div>
              </SortableContext>
            </DndContext>
          </div>
        )}

        {sortedDocuments.length === 0 && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            className="text-center py-12 bg-gray-50 dark:bg-gray-800 rounded-lg border border-gray-200 dark:border-gray-700"
          >
            <div className="w-16 h-16 bg-gray-100 dark:bg-gray-700 rounded-full flex items-center justify-center mx-auto mb-4">
              <DocumentCheckIcon className="w-8 h-8 text-gray-400" />
            </div>
            <h3 className="text-lg font-medium text-gray-900 dark:text-white mb-2">
              No documents selected
            </h3>
            <p className="text-gray-500 dark:text-gray-400">
              Select at least one document from the list above to continue.
            </p>
          </motion.div>
        )}

        {/* Navigation */}
        <div className="flex justify-between items-center pt-8 border-t border-gray-200 dark:border-gray-700">
          <motion.button
            onClick={onPrevious}
            whileHover={{ scale: 1.02 }}
            whileTap={{ scale: 0.98 }}
            className="btn btn-secondary btn-lg"
          >
            <ArrowLeftIcon className="w-5 h-5 mr-2" />
            Previous
          </motion.button>

          <motion.button
            onClick={onNext}
            disabled={!canProceed}
            whileHover={canProceed ? { scale: 1.02 } : {}}
            whileTap={canProceed ? { scale: 0.98 } : {}}
            className={cn(
              "btn btn-primary btn-lg",
              !canProceed && "opacity-50 cursor-not-allowed"
            )}
          >
            Generate Packet
            <ArrowRightIcon className="w-5 h-5 ml-2" />
          </motion.button>
        </div>
      </div>
    </motion.div>
  )
}
