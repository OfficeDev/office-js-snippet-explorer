declare module Word {
    class Body extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_style;
        private m_text;
        private m__ReferenceId;
        contentControls: Word.ContentControlCollection;
        font: Word.Font;
        inlinePictures: Word.InlinePictureCollection;
        paragraphs: Word.ParagraphCollection;
        parentContentControl: Word.ContentControl;
        style: string;
        /**
         *
         * Returns text of the body. Read only.
         */
        text: string;
        _ReferenceId: string;
        clear(): void;
        getHtml(): OfficeExtension.ClientResult<string>;
        getOoxml(): OfficeExtension.ClientResult<string>;
        insertBreak(breakType: string, insertLocation: string): void;
        insertContentControl(): Word.ContentControl;
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        insertHtml(html: string, insertLocation: string): Word.Range;
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        insertText(text: string, insertLocation: string): Word.Range;
        search(searchText: string, searchOptions?: Word.SearchOptions): Word.SearchResultCollection;
        select(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Body;
        _initReferenceId(value: string): void;
    }
    class ContentControl extends OfficeExtension.ClientObject {
        private m_appearance;
        private m_cannotDelete;
        private m_cannotEdit;
        private m_color;
        private m_contentControls;
        private m_font;
        private m_id;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_placeholderText;
        private m_removeWhenEdited;
        private m_style;
        private m_tag;
        private m_text;
        private m_title;
        private m_type;
        private m__ReferenceId;
        contentControls: Word.ContentControlCollection;
        font: Word.Font;
        inlinePictures: Word.InlinePictureCollection;
        paragraphs: Word.ParagraphCollection;
        parentContentControl: Word.ContentControl;
        appearance: string;
        cannotDelete: boolean;
        cannotEdit: boolean;
        color: string;
        id: number;
        placeholderText: string;
        removeWhenEdited: boolean;
        style: string;
        tag: string;
        text: string;
        title: string;
        type: string;
        _ReferenceId: string;
        clear(): void;
        delete(keepContent: boolean): void;
        getHtml(): OfficeExtension.ClientResult<string>;
        getOoxml(): OfficeExtension.ClientResult<string>;
        insertBreak(breakType: string, insertLocation: string): void;
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        insertHtml(html: string, insertLocation: string): Word.Range;
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        insertText(text: string, insertLocation: string): Word.Range;
        search(searchText: string, searchOptions?: Word.SearchOptions): Word.SearchResultCollection;
        select(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.ContentControl;
        _initReferenceId(value: string): void;
    }
    class ContentControlCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.ContentControl>;
        _ReferenceId: string;
        getById(id: number): Word.ContentControl;
        getByTag(tag: string): Word.ContentControlCollection;
        getByTitle(title: string): Word.ContentControlCollection;
        getItem(index: number): Word.ContentControl;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.ContentControlCollection;
        _initReferenceId(value: string): void;
    }
    class Document extends OfficeExtension.ClientObject {
        private m_body;
        private m_contentControls;
        private m_saved;
        private m_sections;
        body: Word.Body;
        contentControls: Word.ContentControlCollection;
        sections: Word.SectionCollection;
        saved: boolean;
        /**
         *
         * Gets the currently selected range from the document.
         *
         */
        getSelection(): Word.Range;
        save(): void;
        _GetObjectByReferenceId(referenceId: string): OfficeExtension.ClientResult<any>;
        _GetObjectTypeNameByReferenceId(referenceId: string): OfficeExtension.ClientResult<string>;
        _RemoveAllReferences(): void;
        _RemoveReference(referenceId: string): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Document;
    }
    class Font extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_doubleStrikeThrough;
        private m_highlightColor;
        private m_italic;
        private m_name;
        private m_size;
        private m_strikeThrough;
        private m_subscript;
        private m_superscript;
        private m_underline;
        private m__ReferenceId;
        bold: boolean;
        color: string;
        doubleStrikeThrough: boolean;
        highlightColor: string;
        italic: boolean;
        name: string;
        size: number;
        strikeThrough: boolean;
        subscript: boolean;
        superscript: boolean;
        underline: string;
        _ReferenceId: string;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Font;
        _initReferenceId(value: string): void;
    }
    class InlinePicture extends OfficeExtension.ClientObject {
        private m_altTextDescription;
        private m_altTextTitle;
        private m_height;
        private m_hyperlink;
        private m_lockAspectRatio;
        private m_parentContentControl;
        private m_width;
        private m__Id;
        private m__ReferenceId;
        parentContentControl: Word.ContentControl;
        altTextDescription: string;
        altTextTitle: string;
        height: number;
        hyperlink: string;
        lockAspectRatio: boolean;
        width: number;
        _Id: number;
        _ReferenceId: string;
        getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        insertContentControl(): Word.ContentControl;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.InlinePicture;
        _initReferenceId(value: string): void;
    }
    class InlinePictureCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.InlinePicture>;
        _ReferenceId: string;
        getItem(index: number): Word.InlinePicture;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.InlinePictureCollection;
        _initReferenceId(value: string): void;
    }
    class Paragraph extends OfficeExtension.ClientObject {
        private m_alignment;
        private m_contentControls;
        private m_firstLineIndent;
        private m_font;
        private m_inlinePictures;
        private m_leftIndent;
        private m_lineSpacing;
        private m_lineUnitAfter;
        private m_lineUnitBefore;
        private m_outlineLevel;
        private m_parentContentControl;
        private m_rightIndent;
        private m_spaceAfter;
        private m_spaceBefore;
        private m_style;
        private m_text;
        private m__Id;
        private m__ReferenceId;
        contentControls: Word.ContentControlCollection;
        font: Word.Font;
        inlinePictures: Word.InlinePictureCollection;
        parentContentControl: Word.ContentControl;
        alignment: string;
        firstLineIndent: number;
        leftIndent: number;
        lineSpacing: number;
        lineUnitAfter: number;
        lineUnitBefore: number;
        outlineLevel: number;
        rightIndent: number;
        spaceAfter: number;
        spaceBefore: number;
        style: string;
        text: string;
        _Id: number;
        _ReferenceId: string;
        clear(): void;
        delete(): void;
        getHtml(): OfficeExtension.ClientResult<string>;
        getOoxml(): OfficeExtension.ClientResult<string>;
        insertBreak(breakType: string, insertLocation: string): void;
        insertContentControl(): Word.ContentControl;
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        insertHtml(html: string, insertLocation: string): Word.Range;
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: string): Word.InlinePicture;
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        insertText(text: string, insertLocation: string): Word.Range;
        search(searchText: string, searchOptions?: Word.SearchOptions): Word.SearchResultCollection;
        select(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Paragraph;
        _initReferenceId(value: string): void;
    }
    class ParagraphCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Paragraph>;
        _ReferenceId: string;
        getItem(index: number): Word.Paragraph;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.ParagraphCollection;
        _initReferenceId(value: string): void;
    }
    class Range extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_paragraphs;
        private m_parentContentControl;
        private m_style;
        private m_text;
        private m__Id;
        private m__ReferenceId;
        contentControls: Word.ContentControlCollection;
        font: Word.Font;
        paragraphs: Word.ParagraphCollection;
        parentContentControl: Word.ContentControl;
        style: string;
        text: string;
        _Id: number;
        _ReferenceId: string;
        clear(): void;
        delete(): void;
        getHtml(): OfficeExtension.ClientResult<string>;
        getOoxml(): OfficeExtension.ClientResult<string>;
        insertBreak(breakType: string, insertLocation: string): void;
        insertContentControl(): Word.ContentControl;
        insertFileFromBase64(base64File: string, insertLocation: string): Word.Range;
        insertHtml(html: string, insertLocation: string): Word.Range;
        insertOoxml(ooxml: string, insertLocation: string): Word.Range;
        insertParagraph(paragraphText: string, insertLocation: string): Word.Paragraph;
        insertText(text: string, insertLocation: string): Word.Range;
        search(searchText: string, searchOptions?: Word.SearchOptions): Word.SearchResultCollection;
        select(): void;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Range;
        _initReferenceId(value: string): void;
    }
    class SearchOptions extends OfficeExtension.ClientObject {
        private m_ignorePunct;
        private m_ignoreSpace;
        private m_matchCase;
        private m_matchPrefix;
        private m_matchSoundsLike;
        private m_matchSuffix;
        private m_matchWholeWord;
        private m_matchWildCards;
        ignorePunct: boolean;
        ignoreSpace: boolean;
        matchCase: boolean;
        matchPrefix: boolean;
        matchSoundsLike: boolean;
        matchSuffix: boolean;
        matchWholeWord: boolean;
        matchWildCards: boolean;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.SearchOptions;
        /**
         * Create a new instance of Word.SearchOptions object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
    }
    class SearchResultCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Range>;
        _ReferenceId: string;
        getItem(index: number): Word.Range;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.SearchResultCollection;
        _initReferenceId(value: string): void;
    }
    class Section extends OfficeExtension.ClientObject {
        private m_body;
        private m__Id;
        private m__ReferenceId;
        body: Word.Body;
        _Id: number;
        _ReferenceId: string;
        getFooter(type: string): Word.Body;
        getHeader(type: string): Word.Body;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.Section;
        _initReferenceId(value: string): void;
    }
    class SectionCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        /** Gets the loaded child items in this collection. */
        items: Array<Word.Section>;
        _ReferenceId: string;
        getItem(index: number): Word.Section;
        _KeepReference(): void;
        /** Handle results returned from the document
         * @private
         */
        _handleResult(value: any): void;
        /**
         * Load properties
         */
        load(option?: string | OfficeExtension.LoadOption): Word.SectionCollection;
        _initReferenceId(value: string): void;
    }
    /**
     *
     * ContentControl types
     */
    module ContentControlType {
        var richText: string;
    }
    /**
     *
     * ContentControl appearance
     */
    module ContentControlAppearance {
        var boundingBox: string;
        var tags: string;
        var hidden: string;
    }
    /**
     *
     * Underline types
     */
    module UnderlineType {
        var none: string;
        var single: string;
        var word: string;
        var double: string;
        var dotted: string;
        var hidden: string;
        var thick: string;
        var dashLine: string;
        var dotLine: string;
        var dotDashLine: string;
        var twoDotDashLine: string;
        var wave: string;
    }
    module BreakType {
        var page: string;
        var column: string;
        var next: string;
        var sectionContinuous: string;
        var sectionEven: string;
        var sectionOdd: string;
        var line: string;
        var lineClearLeft: string;
        var lineClearRight: string;
        var textWrapping: string;
    }
    module InsertLocation {
        var before: string;
        var after: string;
        var start: string;
        var end: string;
        var replace: string;
    }
    module Alignment {
        var unknown: string;
        var left: string;
        var centered: string;
        var right: string;
        var justified: string;
    }
    module HeaderFooterType {
        var primary: string;
        var firstPage: string;
        var evenPages: string;
    }
    module ErrorCodes {
        var accessDenied: string;
        var generalException: string;
        var invalidArgument: string;
        var itemNotFound: string;
        var notImplemented: string;
    }
}
declare module Word {
    class RequestContext extends OfficeExtension.ClientRequestContext {
        private m_document;
        constructor(url?: string);
        document: Document;
    }
}
