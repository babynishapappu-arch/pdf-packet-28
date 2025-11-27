import { PDFDocument, PDFPage, rgb, StandardFonts, PageSizes } from 'pdf-lib'
import { supabase } from './supabaseClient'
import type { ProjectFormData, SelectedDocument } from '@/types'

interface DocumentSection {
  name: string
  type: string
  startPage: number
  pageCount: number
}

class PDFService {
  private async getPdfBytes(url: string): Promise<Uint8Array> {
    const response = await fetch(url)
    if (!response.ok) {
      throw new Error(`Failed to fetch PDF: ${response.statusText}`)
    }
    const arrayBuffer = await response.arrayBuffer()
    return new Uint8Array(arrayBuffer)
  }

  async getDocumentUrl(filePath: string): Promise<string> {
    try {
      const { data, error } = await supabase.storage
        .from('documents')
        .createSignedUrl(filePath, 3600)

      if (error) throw error
      return data.signedUrl
    } catch (error) {
      console.error('Error generating signed URL:', error)
      throw new Error('Failed to generate document URL')
    }
  }

  async generatePacket(
    formData: ProjectFormData,
    selectedDocuments: SelectedDocument[]
  ): Promise<Uint8Array> {
    console.log('Starting PDF generation...')
    const finalPdf = await PDFDocument.create()

    // Sort selected documents by order
    const sortedDocs = selectedDocuments
      .filter(doc => doc.selected)
      .sort((a, b) => a.order - b.order)

    // Get all document names for the submittal form
    const selectedDocumentNames = sortedDocs.map(doc => doc.document.name)

    // 1. Add Cover Page (Submittal Form)
    let submittalFormPageCount = 0
    try {
      console.log('Adding submittal form...')
      await this.addCoverPage(finalPdf, formData, selectedDocumentNames)
      submittalFormPageCount = finalPdf.getPageCount()
      console.log(`Added ${submittalFormPageCount} submittal form pages`)
    } catch (error) {
      console.error('Error creating cover page:', error)
      await this.addErrorPage(finalPdf, 'Cover Page', 'Failed to create cover page')
      submittalFormPageCount = 1
    }

    // 2. Add Product Info Page
    try {
      console.log('Adding product info page...')
      await this.addProductInfoPage(finalPdf, formData)
      console.log('Product info page added')
    } catch (error) {
      console.error('Error adding product info:', error)
      await this.addErrorPage(finalPdf, 'Product Info', 'Failed to add product information')
    }

    const submittalAndProductInfoPageCount = finalPdf.getPageCount()
    const tocPageIndex = submittalAndProductInfoPageCount

    // 3. Add document sections first (without TOC)
    const documentSections: DocumentSection[] = []
    let currentPageNumber = submittalAndProductInfoPageCount + 2 // +2 because TOC will be inserted at position after submittal pages

    for (const doc of sortedDocs) {
      try {
        console.log(`Processing document: ${doc.document.name}`)
        const sectionStartPage = currentPageNumber

        // Add section divider
        await this.addSectionDivider(finalPdf, doc.document.name, doc.document.type)
        currentPageNumber++

        // Get signed URL and download the document
        const signedUrl = await this.getDocumentUrl(doc.document.url)
        console.log(`Fetching PDF from: ${signedUrl}`)

        const pdfBytes = await this.getPdfBytes(signedUrl)
        const sourcePdf = await PDFDocument.load(pdfBytes)
        const pageIndices = sourcePdf.getPageIndices()
        const pages = await finalPdf.copyPages(sourcePdf, pageIndices)
        pages.forEach(page => finalPdf.addPage(page))

        currentPageNumber += pages.length

        documentSections.push({
          name: doc.document.name,
          type: doc.document.type,
          startPage: sectionStartPage,
          pageCount: pages.length + 1
        })

        console.log(`Added ${pages.length} pages from ${doc.document.name}`)
      } catch (error) {
        console.error(`Error processing document ${doc.document.name}:`, error)
        await this.addErrorPage(finalPdf, doc.document.name, 'Failed to process document')
        currentPageNumber++
      }
    }

    // 4. Now insert TOC with actual document sections (only once)
    const tocPageNumber = tocPageIndex + 1
    const tocPage = await this.createTableOfContents(finalPdf, documentSections, tocPageNumber)
    finalPdf.insertPage(tocPageIndex, tocPage)

    // 5. Add page numbers to specific pages only
    await this.addSelectivePageNumbers(
      finalPdf,
      tocPageIndex,
      documentSections
    )

    const pdfBytes = await finalPdf.save()
    console.log(`Packet generated successfully: ${pdfBytes.length} bytes`)

    return pdfBytes
  }

  private async addCoverPage(
    pdf: PDFDocument,
    projectData: ProjectFormData,
    selectedDocumentNames: string[]
  ): Promise<void> {
    let page = pdf.addPage(PageSizes.Letter)
    const { width, height } = page.getSize()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdf.embedFont(StandardFonts.HelveticaBold)

    // Colors
    const nexgenBlue = rgb(0, 0.637, 0.792)
    const darkGray = rgb(0.13, 0.13, 0.13)
    const mediumGray = rgb(0.27, 0.27, 0.27)
    const headerDark = rgb(0.078, 0.078, 0.078)

    // Dark header bar
    page.drawRectangle({
      x: 0,
      y: height - 80,
      width: width,
      height: 80,
      color: headerDark,
    })

    // Logo placeholder
    page.drawText('NEXGEN', {
      x: 50,
      y: height - 45,
      size: 18,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    // Section identifier
    const sectionText = 'SECTION 06 16 26'
    page.drawText(sectionText, {
      x: width - 145,
      y: height - 45,
      size: 10,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    // Title
    const titleColor = rgb(0.094, 0.094, 0.098)
    const isStructuralFloor = projectData.productType === 'structural-floor'

    if (isStructuralFloor) {
      page.drawText('MAXTERRA® MgO Non-Combustible Structural', {
        x: 55,
        y: height - 147,
        size: 18,
        font: font,
        color: titleColor,
      })
      page.drawText('Floor Panels Submittal Form', {
        x: 55,
        y: height - 167,
        size: 18,
        font: font,
        color: titleColor,
      })
    } else {
      page.drawText('MAXTERRA® MgO Non-Combustible', {
        x: 55,
        y: height - 147,
        size: 18,
        font: font,
        color: titleColor,
      })
      page.drawText('Underlayment Panels Submittal Form', {
        x: 55,
        y: height - 167,
        size: 18,
        font: font,
        color: titleColor,
      })
    }

    // Form fields
    let currentY = height - 210
    const labelX = 55
    const valueX = 155
    const fieldHeight = 22
    const fieldWidth = width - valueX - 55
    const fieldSpacing = 4

    const drawFormField = (label: string, value: string, y: number) => {
      page.drawText(label, {
        x: labelX,
        y: y + 6,
        size: 10,
        font: boldFont,
        color: darkGray,
      })

      page.drawRectangle({
        x: valueX,
        y: y,
        width: fieldWidth,
        height: fieldHeight,
        color: rgb(0.95, 0.95, 0.95),
        borderColor: rgb(0.9, 0.9, 0.9),
        borderWidth: 0.5,
      })

      page.drawText(value || '', {
        x: valueX + 10,
        y: y + 6,
        size: 10,
        font: font,
        color: rgb(0, 0, 0),
      })
    }

    drawFormField('Submitted To', projectData.submittedTo, currentY)
    currentY -= (fieldHeight + fieldSpacing)

    drawFormField('Project Name', projectData.projectName, currentY)
    currentY -= (fieldHeight + fieldSpacing)

    drawFormField('Project Number', projectData.projectNumber || '', currentY)
    currentY -= (fieldHeight + fieldSpacing)

    drawFormField('Prepared By', projectData.preparedBy, currentY)
    currentY -= (fieldHeight + fieldSpacing)

    drawFormField('Email Address', projectData.emailAddress, currentY)
    currentY -= (fieldHeight + fieldSpacing)

    drawFormField('Phone Number', projectData.phoneNumber, currentY)
    currentY -= (fieldHeight + fieldSpacing)

    drawFormField('Date', projectData.date, currentY)
    currentY -= (fieldHeight + 15)

    // Status/Action checkboxes
    page.drawText('Status / Action', {
      x: labelX,
      y: currentY,
      size: 10,
      font: boldFont,
      color: darkGray,
    })
    currentY -= 10

    const checkboxSize = 12
    const checkboxSpacing = 130
    let checkboxX = valueX

    const drawCheckbox = (label: string, checked: boolean, x: number, y: number) => {
      page.drawRectangle({
        x: x,
        y: y,
        width: checkboxSize,
        height: checkboxSize,
        color: rgb(0.95, 0.95, 0.95),
        borderColor: rgb(0.9, 0.9, 0.9),
        borderWidth: 0.5,
      })

      if (checked) {
        page.drawText('X', {
          x: x + 3,
          y: y + 2,
          size: 9,
          font: boldFont,
          color: nexgenBlue,
        })
      }

      page.drawText(label, {
        x: x + checkboxSize + 5,
        y: y + 2,
        size: 10,
        font: font,
        color: rgb(0, 0, 0),
      })
    }

    drawCheckbox('For Review', projectData.status.forReview, checkboxX, currentY + 3)
    drawCheckbox('For Approval', projectData.status.forApproval, checkboxX + checkboxSpacing, currentY + 3)
    currentY -= 18
    drawCheckbox('For Record', projectData.status.forRecord, checkboxX, currentY + 3)
    drawCheckbox('For Information Only', projectData.status.forInformationOnly, checkboxX + checkboxSpacing, currentY + 3)

    currentY -= 30

    // Submittal Type section
    page.drawText('Submittal Type (check all that apply):', {
      x: labelX,
      y: currentY,
      size: 15,
      font: boldFont,
      color: darkGray,
    })
    currentY -= 20

    // Draw selected documents
    const minYForContent = 100
    let currentPage = page
    const checkboxLineSpacing = 14

    for (let i = 0; i < selectedDocumentNames.length; i++) {
      const docName = selectedDocumentNames[i]

      if (currentY < minYForContent) {
        currentPage = pdf.addPage(PageSizes.Letter)
        currentY = height - 140

        // Add header to new page
        currentPage.drawRectangle({
          x: 0,
          y: height - 80,
          width: width,
          height: 80,
          color: headerDark,
        })

        currentPage.drawText('NEXGEN', {
          x: 50,
          y: height - 45,
          size: 18,
          font: boldFont,
          color: rgb(1, 1, 1),
        })
      }

      const checkboxY = currentY
      currentPage.drawRectangle({
        x: labelX,
        y: checkboxY,
        width: checkboxSize,
        height: checkboxSize,
        color: rgb(0.95, 0.95, 0.95),
        borderColor: rgb(0.9, 0.9, 0.9),
        borderWidth: 0.5,
      })

      currentPage.drawText('X', {
        x: labelX + 3,
        y: checkboxY + 2,
        size: 9,
        font: boldFont,
        color: nexgenBlue,
      })

      currentPage.drawText(docName, {
        x: labelX + checkboxSize + 5,
        y: checkboxY + 2,
        size: 10,
        font: font,
        color: rgb(0, 0, 0),
      })

      currentY -= checkboxLineSpacing
    }

    currentY -= 10

    // Product section
    if (currentY < minYForContent + 100) {
      currentPage = pdf.addPage(PageSizes.Letter)
      currentY = height - 140

      currentPage.drawRectangle({
        x: 0,
        y: height - 80,
        width: width,
        height: 80,
        color: headerDark,
      })

      currentPage.drawText('NEXGEN', {
        x: 50,
        y: height - 45,
        size: 18,
        font: boldFont,
        color: rgb(1, 1, 1),
      })
    }

    currentPage.drawText('Product:', {
      x: labelX,
      y: currentY,
      size: 10,
      font: boldFont,
      color: darkGray,
    })

    const productText = isStructuralFloor
      ? 'MAXTERRA® MgO Non-Combustible Structural Floor Panels'
      : 'MAXTERRA® MgO Fire- And Water-Resistant Underlayment Panels'

    currentPage.drawText(productText, {
      x: valueX,
      y: currentY,
      size: 9,
      font: font,
      color: darkGray,
    })

    // Footer
    const footerY = 120
    currentPage.drawText('NEXGEN® Building Products, LLC', {
      x: labelX,
      y: footerY,
      size: 9,
      font: boldFont,
      color: darkGray,
    })
    currentPage.drawText('1504 Manhattan Ave West, #300 Brandon, FL 34205', {
      x: labelX,
      y: footerY - 12,
      size: 8,
      font: font,
      color: mediumGray,
    })
    currentPage.drawText('(727) 634-5534', {
      x: labelX,
      y: footerY - 24,
      size: 8,
      font: font,
      color: mediumGray,
    })
    currentPage.drawText('Technical Support: support@nexgenbp.com', {
      x: labelX,
      y: footerY - 36,
      size: 8,
      font: font,
      color: mediumGray,
    })

    const versionText = 'Version 1.0 October 2025 © 2025 NEXGEN Building Products'
    const versionWidth = font.widthOfTextAtSize(versionText, 7)
    currentPage.drawText(versionText, {
      x: width - versionWidth - 50,
      y: 50,
      size: 7,
      font: font,
      color: mediumGray,
    })
  }

  private async addProductInfoPage(pdf: PDFDocument, projectData: ProjectFormData): Promise<void> {
    const page = pdf.addPage(PageSizes.Letter)
    const { width, height } = page.getSize()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdf.embedFont(StandardFonts.HelveticaBold)

    const isStructuralFloor = projectData.productType === 'structural-floor'
    const headerDark = rgb(0.078, 0.078, 0.078)

    // Dark header bar
    page.drawRectangle({
      x: 0,
      y: height - 80,
      width: width,
      height: 80,
      color: headerDark,
    })

    page.drawText('NEXGEN', {
      x: 50,
      y: height - 45,
      size: 18,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    const sectionText = isStructuralFloor ? 'SECTION 06 16 23' : 'SECTION 06 16 26'
    page.drawText(sectionText, {
      x: width - 145,
      y: height - 45,
      size: 10,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    // Add content based on product type
    const margin = 50
    let currentY = height - 100

    page.drawText('Product Information', {
      x: margin,
      y: currentY,
      size: 16,
      font: boldFont,
      color: rgb(0.302, 0.298, 0.298),
    })

    currentY -= 30

    const productName = isStructuralFloor
      ? 'MAXTERRA® MgO Non-Combustible Structural Floor Panels'
      : 'MAXTERRA® MgO Fire- And Water-Resistant Underlayment Panels'

    page.drawText(productName, {
      x: margin,
      y: currentY,
      size: 12,
      font: boldFont,
      color: rgb(0, 0, 0),
    })
  }

  private async createTableOfContents(
    pdf: PDFDocument,
    sections: DocumentSection[],
    tocPageNumber: number
  ): Promise<PDFPage> {
    const page = pdf.addPage(PageSizes.Letter)
    const { width, height } = page.getSize()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdf.embedFont(StandardFonts.HelveticaBold)
    const headerDark = rgb(0.078, 0.078, 0.078)

    // Dark header bar
    page.drawRectangle({
      x: 0,
      y: height - 80,
      width: width,
      height: 80,
      color: headerDark,
    })

    page.drawText('NEXGEN', {
      x: 50,
      y: height - 45,
      size: 18,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    page.drawText('Table of Contents', {
      x: width - 180,
      y: height - 45,
      size: 11,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    let currentY = height - 110
    const margin = 50
    const lineHeight = 25

    page.drawText('Table of Contents', {
      x: margin,
      y: currentY,
      size: 18,
      font: boldFont,
      color: rgb(0.094, 0.094, 0.098),
    })
    currentY -= 30

    sections.forEach((section) => {
      if (currentY < 100) return

      page.drawText(section.name, {
        x: margin,
        y: currentY,
        size: 11,
        font: font,
        color: rgb(0, 0, 0),
      })

      page.drawText(`Page ${section.startPage}`, {
        x: width - 90,
        y: currentY,
        size: 11,
        font: boldFont,
        color: rgb(0, 0, 0),
      })

      currentY -= lineHeight
    })

    return page
  }

  private async addSectionDivider(
    pdf: PDFDocument,
    documentName: string,
    documentType: string
  ): Promise<void> {
    const page = pdf.addPage(PageSizes.Letter)
    const { width, height } = page.getSize()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdf.embedFont(StandardFonts.HelveticaBold)
    const headerDark = rgb(0.078, 0.078, 0.078)

    page.drawRectangle({
      x: 0,
      y: height - 80,
      width: width,
      height: 80,
      color: headerDark,
    })

    page.drawText('NEXGEN', {
      x: 50,
      y: height - 45,
      size: 18,
      font: boldFont,
      color: rgb(1, 1, 1),
    })

    page.drawText('Document Section', {
      x: width - 180,
      y: height - 45,
      size: 11,
      font: font,
      color: rgb(1, 1, 1),
    })

    const centerY = height / 2
    const nameSize = 40
    const nameWidth = boldFont.widthOfTextAtSize(documentName, nameSize)

    page.drawText(documentName, {
      x: (width - nameWidth) / 2,
      y: centerY,
      size: nameSize,
      font: boldFont,
      color: rgb(0.13, 0.13, 0.13),
    })

    const lineWidth = 200
    page.drawLine({
      start: { x: (width - lineWidth) / 2, y: centerY - 30 },
      end: { x: (width + lineWidth) / 2, y: centerY - 30 },
      thickness: 2,
      color: rgb(0, 0.637, 0.792),
    })
  }

  private async addErrorPage(
    pdf: PDFDocument,
    documentName: string,
    errorMessage: string
  ): Promise<void> {
    const page = pdf.addPage(PageSizes.Letter)
    const { width, height } = page.getSize()
    const font = await pdf.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdf.embedFont(StandardFonts.HelveticaBold)

    page.drawText('DOCUMENT ERROR', {
      x: 50,
      y: height - 100,
      size: 16,
      font: boldFont,
      color: rgb(0.8, 0.2, 0.2),
    })

    page.drawText(documentName, {
      x: 50,
      y: height - 150,
      size: 14,
      font: boldFont,
      color: rgb(0, 0, 0),
    })

    page.drawText(`Error: ${errorMessage}`, {
      x: 50,
      y: height - 180,
      size: 12,
      font: font,
      color: rgb(0.6, 0.2, 0.2),
    })

    page.drawText('Please contact support if this error persists.', {
      x: 50,
      y: height - 220,
      size: 10,
      font: font,
      color: rgb(0.4, 0.4, 0.4),
    })
  }

  private async addSelectivePageNumbers(
    pdf: PDFDocument,
    submittalAndProductInfoPageCount: number,
    sections: DocumentSection[]
  ): Promise<void> {
    const pages = pdf.getPages()
    const font = await pdf.embedFont(StandardFonts.Helvetica)

    let globalPageNumber = 1

    // Add page numbers to Submittal Form + Product Info pages
    for (let i = 0; i < submittalAndProductInfoPageCount && i < pages.length; i++) {
      const page = pages[i]
      const { width } = page.getSize()
      page.drawText(`${globalPageNumber}`, {
        x: width - 50,
        y: 30,
        size: 10,
        font: font,
        color: rgb(0.4, 0.4, 0.4),
      })
      globalPageNumber++
    }

    // Add page number to Table of Contents
    if (pages.length > submittalAndProductInfoPageCount) {
      const tocPage = pages[submittalAndProductInfoPageCount]
      const { width } = tocPage.getSize()
      tocPage.drawText(`${globalPageNumber}`, {
        x: width - 50,
        y: 30,
        size: 10,
        font: font,
        color: rgb(0.4, 0.4, 0.4),
      })
      globalPageNumber++
    }

    // Add page numbers to Section Dividers only
    let currentIndex = submittalAndProductInfoPageCount + 1
    sections.forEach(section => {
      if (currentIndex < pages.length) {
        const dividerPage = pages[currentIndex]
        const { width } = dividerPage.getSize()
        dividerPage.drawText(`${globalPageNumber}`, {
          x: width - 50,
          y: 30,
          size: 10,
          font: font,
          color: rgb(0.4, 0.4, 0.4),
        })
        globalPageNumber++

        // Skip the actual document pages
        currentIndex += section.pageCount
      }
    })
  }

  async downloadPDF(pdfBytes: Uint8Array, filename: string): Promise<void> {
    const blob = new Blob([pdfBytes], { type: 'application/pdf' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename || 'document-packet.pdf'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  async previewPDF(pdfBytes: Uint8Array): Promise<string> {
    const blob = new Blob([pdfBytes], { type: 'application/pdf' })
    return URL.createObjectURL(blob)
  }
}

export const pdfService = new PDFService()
