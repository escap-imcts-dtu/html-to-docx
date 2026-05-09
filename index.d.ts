/// <reference types="node" />

declare namespace HTMLtoDOCX {
    interface Margins {
        top?: number;
        right?: number;
        bottom?: number;
        left?: number;
        header?: number;
        footer?: number;
        gutter?: number;
    }

    interface PageSize {
        width?: number;
        height?: number;
    }

    interface Row {
        cantSplit?: boolean;
    }

    interface Table {
        row?: Row;
        borderOptions?: {
            size?: number;
            stroke?: string;
            color?: string;
        };
        addSpacingAfter?: boolean;
    }

    interface LineNumberOptions {
        start: number;
        countBy: number;
        restart: "continuous" | "newPage" | "newSection";
    }

    interface HeadingSpacing {
        before?: number;
        after?: number;
    }

    interface HeadingStyle {
        font?: string;
        fontSize?: number;
        bold?: boolean;
        spacing?: HeadingSpacing;
        keepLines?: boolean;
        keepNext?: boolean;
        outlineLevel?: number;
    }

    interface HeadingConfig {
        heading1?: HeadingStyle;
        heading2?: HeadingStyle;
        heading3?: HeadingStyle;
        heading4?: HeadingStyle;
        heading5?: HeadingStyle;
        heading6?: HeadingStyle;
    }

    interface DocumentOptions {
        orientation?: "portrait" | "landscape";
        pageSize?: PageSize;
        margins?: Margins;
        title?: string;
        subject?: string;
        creator?: string;
        keywords?: string[];
        description?: string;
        lastModifiedBy?: string;
        revision?: number;
        createdAt?: Date;
        modifiedAt?: Date;
        headerType?: "default" | "first" | "even";
        header?: boolean;
        footerType?: "default" | "first" | "even";
        footer?: boolean;
        font?: string;
        fontSize?: number;
        complexScriptFontSize?: number;
        table?: Table;
        pageNumber?: boolean;
        skipFirstHeaderFooter?: boolean;
        lineNumber?: boolean;
        lineNumberOptions?: LineNumberOptions;
        numbering?: {
            defaultOrderedListStyleType?: string;
        };
        heading?: HeadingConfig;
        decodeUnicode?: boolean;
        lang?: string;
        direction?: "ltr" | "rtl";
        preprocessing?: {
            skipHTMLMinify?: boolean;
        };
        imageProcessing?: {
            maxRetries?: number;
            verboseLogging?: boolean;
            downloadTimeout?: number;
            maxImageSize?: number;
            retryDelayBase?: number;
            minTimeout?: number;
            maxTimeout?: number;
            minImageSize?: number;
            maxCacheSize?: number;
            maxCacheEntries?: number;
            svgHandling?: "convert" | "native" | "auto";
            suppressSharpWarning?: boolean;
        };
    }
}

declare namespace HTMLtoDOCX {
    interface OoxmlFragmentMedia {
        nameInMedia: string;
        contentType: string;
        data: Buffer;
    }

    interface OoxmlFragmentRelationship {
        relationshipId: number;
        type: string;
        target: string;
        targetMode: string;
    }

    interface OoxmlFragmentNumbering {
        numberingId: number;
        type: 'ol' | 'ul';
        properties: Record<string, unknown>;
    }

    interface OoxmlFragment {
        bodyXml: string;
        media: OoxmlFragmentMedia[];
        relationships: OoxmlFragmentRelationship[];
        numbering: OoxmlFragmentNumbering[];
    }
}

declare function HTMLtoDOCX(
    htmlString: string,
    headerHTMLstring?: string | null,
    documentOptions?: HTMLtoDOCX.DocumentOptions,
    footerHtmlString?: string | null,
): Promise<ArrayBuffer | Blob | Buffer>;

declare namespace HTMLtoDOCX {
    function htmlToOoxmlFragment(
        htmlString: string,
        documentOptions?: HTMLtoDOCX.DocumentOptions,
    ): Promise<HTMLtoDOCX.OoxmlFragment>;
}

export = HTMLtoDOCX;
