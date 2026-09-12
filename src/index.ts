export { ingestTemplate } from "./ingest.js";
export { Presentation } from "./presentation.js";
export { md } from "./rich-text.js";
export { CustomSlide } from "./custom-slide.js";
export { C, FONTS, LAYOUT, LOGO_FILES, applyDesign, currentDesign } from "./design.js";
export type { ColorName, DesignPatch, FontRole, LayoutSpec, LogoRole } from "./design.js";
export { parseDesignFile, readDesignFileSync, serializeDesign } from "./design-loader.js";
export {
  WorkspaceNotFoundError,
  installDir,
  resolveWorkspaceSync,
  tryResolveWorkspaceSync
} from "./workspace.js";
export type { Workspace } from "./workspace.js";
export { createCustomSlideHelpers } from "./custom-slide-helpers.js";
export { FigureRenderer, fitBox } from "./figure.js";
export type { Figure, FigureFit, FigureResult, FigureViewport } from "./figure.js";
export { findChrome } from "./html-shot.js";
export type { ShotFn, ShotRequest } from "./html-shot.js";
export type {
  AddSlideOptions,
  BuildReport,
  BuildReportSlide,
  BuildWarning,
  CustomDeckSlide,
  DeckSlide,
  RenderOptions,
  SlideOverride,
  FigureRecord,
  SlideVariables,
  TemplateDeckSlide,
  TemplateField
} from "./types.js";
export type { CustomSlideContext, CustomSlideOptions } from "./custom-slide.js";
export type { Box, CustomSlideHelpers, Point } from "./custom-slide-helpers.js";
