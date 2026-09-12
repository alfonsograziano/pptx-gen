export type PrimitiveRichText = string | MarkdownText;

export type MarkdownText = {
  kind: "markdown";
  value: string;
};

export type TemplateIndex = Record<string, {
  slideIndex: number;
  metadata: string;
}>;

export type TemplateField = {
  id: string;
  type: "text" | "image";
  shapeId: string;
  name: string;
  originalText: string;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  preserveStyleByDefault: boolean;
  /**
   * Marks a field the deck owns rather than the author. A `page-number` field is
   * rewritten as a live PowerPoint slide-number field at build time, so it is
   * never listed as a fillable variable.
   */
  role?: "page-number";
};

export type FieldsFile = {
  templateId: string;
  librarySlide: number;
  fields: TemplateField[];
};

export type TemplateMetadata = {
  id: string;
  name: string;
  kind: "cloned-slide";
  status: "draft" | "ready";
  version: string;
  source: {
    deck: string;
    slide: number;
    importedOn: string;
  };
  fonts: string[];
  variables: string[];
  tags: string[];
};

export type SlideVariables = Record<string, PrimitiveRichText>;

export type SlideOverride =
  | { op: "delete"; target: string }
  | { op: "hide"; target: string }
  | { op: "move"; target: string; x: number; y: number }
  | { op: "resize"; target: string; w: number; h: number }
  | { op: "styleText"; target: string; fontSize?: number; color?: string; fontFace?: string }
  | { op: "addText"; id: string; text: PrimitiveRichText; x: number; y: number; w: number; h: number; style?: TextStyle }
  | { op: "addSvg"; id: string; path: string; x: number; y: number; w: number; h: number }
  | { op: "addIcon"; id: string; icon: string; x: number; y: number; w: number; h: number; color?: string }
  | { op: "replaceImage"; target: string; path: string };

export type TextStyle = {
  fontFace?: string;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
};

export type AddSlideOptions = {
  templateName: string;
  variables?: SlideVariables;
  overrides?: SlideOverride[];
  /**
   * Optional variant-group id. Slides sharing a group are alternative
   * renderings of the same concept, listed together in the build report so the
   * reviewer can pick one. Variants of a group must be added consecutively.
   */
  group?: string;
};

export type TemplateDeckSlide = {
  kind: "template";
  options: AddSlideOptions;
};

export type CustomDeckSlide = {
  kind: "custom";
  slide: import("./custom-slide.js").CustomSlide;
};

export type DeckSlide = TemplateDeckSlide | CustomDeckSlide;

export type RenderOptions = {
  output: string;
  report?: string;
  screenshots?: string;
  /** Show live, timed build progress in the terminal. Defaults to true. */
  progress?: boolean;
};

export type BuildWarning = {
  code: string;
  message: string;
  slide?: number;
  target?: string;
};

export type BuildReportSlide = {
  /** 1-based position in the output deck. */
  index: number;
  kind: "template" | "custom";
  /** Template id for template slides, the CustomSlide name for custom slides. */
  name: string;
  /** Variant-group id, when this slide is one of several variants of a concept. */
  group?: string;
  /** 1-based position within the group. Set only when `group` is set. */
  variant?: number;
  /** Total slides in this group. Set only when `group` is set. */
  variantCount?: number;
};

export type BuildReport = {
  generatedAt: string;
  output: string;
  templatesUsed: string[];
  customSlidesUsed: string[];
  /** Every slide in output order, with variant-group membership. */
  slides: BuildReportSlide[];
  slidesBuilt: number;
  warnings: BuildWarning[];
  screenshots: string[];
};
