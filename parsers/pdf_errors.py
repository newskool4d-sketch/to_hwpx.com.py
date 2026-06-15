from __future__ import annotations


class PdfParseError(RuntimeError):
    pass


class OpendataloaderPdfError(PdfParseError):
    pass


class OpendataloaderPdfUnavailableError(OpendataloaderPdfError):
    pass


class OpendataloaderPdfApiError(OpendataloaderPdfError):
    pass


class OpendataloaderPdfConversionError(OpendataloaderPdfError):
    pass


class OpendataloaderPdfOutputError(OpendataloaderPdfError):
    pass


class PdfTextFallbackError(PdfParseError):
    pass


class PdfTextExtractionError(PdfParseError):
    pass


class EmptyPdfTextError(PdfTextExtractionError):
    pass
