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

export type BuildReport = {
  generatedAt: string;
  output: string;
  templatesUsed: string[];
  customSlidesUsed: string[];
  slidesBuilt: number;
  warnings: BuildWarning[];
  screenshots: string[];
};
