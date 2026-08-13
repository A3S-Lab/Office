import { type CommandProps, mergeAttributes, Node } from '@tiptap/core';
import type { DOMOutputSpec } from '@tiptap/pm/model';

export type WorkDocumentEquationDisplay = 'inline' | 'block';
export type WorkDocumentEquationJustification =
  | 'left'
  | 'right'
  | 'center'
  | 'centerGroup';
export type WorkDocumentEquationBarPosition = 'top' | 'bottom';
export type WorkDocumentEquationFractionType =
  | 'bar'
  | 'noBar'
  | 'skewed'
  | 'linear';
export type WorkDocumentEquationLimitLocation = 'underOver' | 'subSup';
export type WorkDocumentEquationDelimiterShape = 'centered' | 'match';
export type WorkDocumentEquationMatrixAlignment = 'left' | 'center' | 'right';
export type WorkDocumentEquationMatrixBaseAlignment =
  | 'top'
  | 'center'
  | 'bottom';
export type WorkDocumentEquationSpacingRule =
  | 'single'
  | 'oneAndHalf'
  | 'double'
  | 'exact'
  | 'multiple';
export type WorkDocumentEquationRowSpacingRule =
  WorkDocumentEquationSpacingRule;

export interface WorkDocumentEquationMatrixSpacing {
  rowSpacingRule: WorkDocumentEquationSpacingRule;
  rowSpacing: number;
  columnGapRule: WorkDocumentEquationSpacingRule;
  columnGap: number;
  minimumColumnWidthTwips: number;
}
export type WorkDocumentEquationRunScript =
  | 'roman'
  | 'sansSerif'
  | 'monospace'
  | 'fraktur'
  | 'doubleStruck'
  | 'script';
export type WorkDocumentEquationRunStyle =
  | 'plain'
  | 'italic'
  | 'bold'
  | 'boldItalic';
export type WorkDocumentEquationThemeFont =
  | 'majorEastAsia'
  | 'majorBidi'
  | 'majorAscii'
  | 'majorHAnsi'
  | 'minorEastAsia'
  | 'minorBidi'
  | 'minorAscii'
  | 'minorHAnsi';
export type WorkDocumentEquationThemeColor =
  | 'dark1'
  | 'light1'
  | 'dark2'
  | 'light2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hyperlink'
  | 'followedHyperlink'
  | 'none'
  | 'background1'
  | 'text1'
  | 'background2'
  | 'text2';
export type WorkDocumentEquationWordHighlight =
  | 'black'
  | 'blue'
  | 'cyan'
  | 'green'
  | 'magenta'
  | 'red'
  | 'yellow'
  | 'white'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'darkGray'
  | 'lightGray'
  | 'none';
export type WorkDocumentEquationWordShadingPattern =
  | 'nil'
  | 'clear'
  | 'solid'
  | 'horzStripe'
  | 'vertStripe'
  | 'reverseDiagStripe'
  | 'diagStripe'
  | 'horzCross'
  | 'diagCross'
  | 'thinHorzStripe'
  | 'thinVertStripe'
  | 'thinReverseDiagStripe'
  | 'thinDiagStripe'
  | 'thinHorzCross'
  | 'thinDiagCross'
  | 'pct5'
  | 'pct10'
  | 'pct12'
  | 'pct15'
  | 'pct20'
  | 'pct25'
  | 'pct30'
  | 'pct35'
  | 'pct37'
  | 'pct40'
  | 'pct45'
  | 'pct50'
  | 'pct55'
  | 'pct60'
  | 'pct62'
  | 'pct65'
  | 'pct70'
  | 'pct75'
  | 'pct80'
  | 'pct85'
  | 'pct87'
  | 'pct90'
  | 'pct95';
export type WorkDocumentEquationUnderlineStyle =
  | 'none'
  | 'words'
  | 'single'
  | 'double'
  | 'thick'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashedHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dashDotHeavy'
  | 'dotDotDash'
  | 'dashDotDotHeavy'
  | 'wave'
  | 'wavyHeavy'
  | 'wavyDouble';
export type WorkDocumentEquationWordTextEffect =
  | 'blinkBackground'
  | 'lights'
  | 'antsBlack'
  | 'antsRed'
  | 'shimmer'
  | 'sparkle'
  | 'none';
export type WorkDocumentEquationWordLineBorderStyle =
  | 'nil'
  | 'none'
  | 'single'
  | 'thick'
  | 'double'
  | 'dotted'
  | 'dashed'
  | 'dotDash'
  | 'dotDotDash'
  | 'triple'
  | 'thinThickSmallGap'
  | 'thickThinSmallGap'
  | 'thinThickThinSmallGap'
  | 'thinThickMediumGap'
  | 'thickThinMediumGap'
  | 'thinThickThinMediumGap'
  | 'thinThickLargeGap'
  | 'thickThinLargeGap'
  | 'thinThickThinLargeGap'
  | 'wave'
  | 'doubleWave'
  | 'dashSmallGap'
  | 'dashDotStroked'
  | 'threeDEmboss'
  | 'threeDEngrave'
  | 'outset'
  | 'inset';

export interface WorkDocumentEquationWordRunFonts {
  ascii?: string;
  highAnsi?: string;
  eastAsia?: string;
  complexScript?: string;
  asciiTheme?: WorkDocumentEquationThemeFont;
  highAnsiTheme?: WorkDocumentEquationThemeFont;
  eastAsiaTheme?: WorkDocumentEquationThemeFont;
  complexScriptTheme?: WorkDocumentEquationThemeFont;
  hint?: 'default' | 'eastAsia' | 'cs';
}

export interface WorkDocumentEquationWordColor {
  value?: 'auto' | string;
  theme?: WorkDocumentEquationThemeColor;
  tint?: string;
  shade?: string;
}

export type WorkDocumentEquationWordEffectSchemeColor =
  | Exclude<WorkDocumentEquationThemeColor, 'none'>
  | 'placeholder';

export type WorkDocumentEquationWordColorTransformType =
  | 'tint'
  | 'shade'
  | 'alpha'
  | 'hueMod'
  | 'saturation'
  | 'saturationOffset'
  | 'saturationModulation'
  | 'luminance'
  | 'luminanceOffset'
  | 'luminanceModulation';

export interface WorkDocumentEquationWordColorTransform {
  type: WorkDocumentEquationWordColorTransformType;
  value: number;
}

export type WorkDocumentEquationWordEffectColor =
  | {
      type: 'rgb';
      value: string;
      transforms?: WorkDocumentEquationWordColorTransform[];
    }
  | {
      type: 'scheme';
      value: WorkDocumentEquationWordEffectSchemeColor;
      transforms?: WorkDocumentEquationWordColorTransform[];
    };

export interface WorkDocumentEquationWordGlow {
  radiusEmus?: number;
  color: WorkDocumentEquationWordEffectColor;
}

export type WorkDocumentEquationWordRectangleAlignment =
  | 'none'
  | 'topLeft'
  | 'top'
  | 'topRight'
  | 'left'
  | 'center'
  | 'right'
  | 'bottomLeft'
  | 'bottom'
  | 'bottomRight';

export interface WorkDocumentEquationWordShadowEffect {
  blurRadiusEmus?: number;
  distanceEmus?: number;
  directionDegrees?: number;
  horizontalScalePercent?: number;
  verticalScalePercent?: number;
  horizontalSkewDegrees?: number;
  verticalSkewDegrees?: number;
  alignment?: WorkDocumentEquationWordRectangleAlignment;
  color: WorkDocumentEquationWordEffectColor;
}

export interface WorkDocumentEquationWordReflectionEffect {
  blurRadiusEmus?: number;
  startOpacityPercent?: number;
  startPositionPercent?: number;
  endOpacityPercent?: number;
  endPositionPercent?: number;
  distanceEmus?: number;
  directionDegrees?: number;
  fadeDirectionDegrees?: number;
  horizontalScalePercent?: number;
  verticalScalePercent?: number;
  horizontalSkewDegrees?: number;
  verticalSkewDegrees?: number;
  alignment?: WorkDocumentEquationWordRectangleAlignment;
}

export type WorkDocumentEquationWordTextOutlineCap =
  | 'round'
  | 'square'
  | 'flat';

export type WorkDocumentEquationWordTextOutlineCompound =
  | 'single'
  | 'double'
  | 'thickThin'
  | 'thinThick'
  | 'triple';

export type WorkDocumentEquationWordTextOutlineAlignment = 'center' | 'inset';

export type WorkDocumentEquationWordPresetLineDash =
  | 'solid'
  | 'dot'
  | 'systemDot'
  | 'dash'
  | 'systemDash'
  | 'longDash'
  | 'dashDot'
  | 'systemDashDot'
  | 'longDashDot'
  | 'longDashDotDot'
  | 'systemDashDotDot';

export type WorkDocumentEquationWordGradientPath =
  | 'shape'
  | 'circle'
  | 'rectangle';

export interface WorkDocumentEquationWordGradientStop {
  positionPercent: number;
  color: WorkDocumentEquationWordEffectColor;
}

export interface WorkDocumentEquationWordGradientFillRectangle {
  leftPercent?: number;
  topPercent?: number;
  rightPercent?: number;
  bottomPercent?: number;
}

export type WorkDocumentEquationWordGradientShade =
  | {
      type: 'linear';
      angleDegrees?: number;
      scaled?: boolean;
    }
  | {
      type: 'path';
      path?: WorkDocumentEquationWordGradientPath;
      fillToRectangle?: WorkDocumentEquationWordGradientFillRectangle;
    };

export type WorkDocumentEquationWordEffectFill =
  | { type: 'none' }
  | { type: 'solid'; color?: WorkDocumentEquationWordEffectColor }
  | {
      type: 'gradient';
      stops?: WorkDocumentEquationWordGradientStop[];
      shade?: WorkDocumentEquationWordGradientShade;
    };

export interface WorkDocumentEquationWordTextFillEffect {
  fill?: WorkDocumentEquationWordEffectFill;
}

export type WorkDocumentEquationWordPresetCamera =
  | 'legacyObliqueTopLeft'
  | 'legacyObliqueTop'
  | 'legacyObliqueTopRight'
  | 'legacyObliqueLeft'
  | 'legacyObliqueFront'
  | 'legacyObliqueRight'
  | 'legacyObliqueBottomLeft'
  | 'legacyObliqueBottom'
  | 'legacyObliqueBottomRight'
  | 'legacyPerspectiveTopLeft'
  | 'legacyPerspectiveTop'
  | 'legacyPerspectiveTopRight'
  | 'legacyPerspectiveLeft'
  | 'legacyPerspectiveFront'
  | 'legacyPerspectiveRight'
  | 'legacyPerspectiveBottomLeft'
  | 'legacyPerspectiveBottom'
  | 'legacyPerspectiveBottomRight'
  | 'orthographicFront'
  | 'isometricTopUp'
  | 'isometricTopDown'
  | 'isometricBottomUp'
  | 'isometricBottomDown'
  | 'isometricLeftUp'
  | 'isometricLeftDown'
  | 'isometricRightUp'
  | 'isometricRightDown'
  | 'isometricOffAxis1Left'
  | 'isometricOffAxis1Right'
  | 'isometricOffAxis1Top'
  | 'isometricOffAxis2Left'
  | 'isometricOffAxis2Right'
  | 'isometricOffAxis2Top'
  | 'isometricOffAxis3Left'
  | 'isometricOffAxis3Right'
  | 'isometricOffAxis3Bottom'
  | 'isometricOffAxis4Left'
  | 'isometricOffAxis4Right'
  | 'isometricOffAxis4Bottom'
  | 'obliqueTopLeft'
  | 'obliqueTop'
  | 'obliqueTopRight'
  | 'obliqueLeft'
  | 'obliqueRight'
  | 'obliqueBottomLeft'
  | 'obliqueBottom'
  | 'obliqueBottomRight'
  | 'perspectiveFront'
  | 'perspectiveLeft'
  | 'perspectiveRight'
  | 'perspectiveAbove'
  | 'perspectiveBelow'
  | 'perspectiveAboveLeftFacing'
  | 'perspectiveAboveRightFacing'
  | 'perspectiveContrastingLeftFacing'
  | 'perspectiveContrastingRightFacing'
  | 'perspectiveHeroicLeftFacing'
  | 'perspectiveHeroicRightFacing'
  | 'perspectiveHeroicExtremeLeftFacing'
  | 'perspectiveHeroicExtremeRightFacing'
  | 'perspectiveRelaxed'
  | 'perspectiveRelaxedModerately';

export type WorkDocumentEquationWordLightRigPreset =
  | 'legacyFlat1'
  | 'legacyFlat2'
  | 'legacyFlat3'
  | 'legacyFlat4'
  | 'legacyNormal1'
  | 'legacyNormal2'
  | 'legacyNormal3'
  | 'legacyNormal4'
  | 'legacyHarsh1'
  | 'legacyHarsh2'
  | 'legacyHarsh3'
  | 'legacyHarsh4'
  | 'threePoint'
  | 'balanced'
  | 'soft'
  | 'harsh'
  | 'flood'
  | 'contrasting'
  | 'morning'
  | 'sunrise'
  | 'sunset'
  | 'chilly'
  | 'freezing'
  | 'flat'
  | 'twoPoint'
  | 'glow'
  | 'brightRoom';

export type WorkDocumentEquationWordLightRigDirection =
  | 'topLeft'
  | 'top'
  | 'topRight'
  | 'left'
  | 'right'
  | 'bottomLeft'
  | 'bottom'
  | 'bottomRight';

export interface WorkDocumentEquationWordScene3DRotation {
  latitudeDegrees: number;
  longitudeDegrees: number;
  revolutionDegrees: number;
}

export interface WorkDocumentEquationWordScene3DLightRig {
  preset: WorkDocumentEquationWordLightRigPreset;
  direction: WorkDocumentEquationWordLightRigDirection;
  rotation?: WorkDocumentEquationWordScene3DRotation;
}

export interface WorkDocumentEquationWordScene3D {
  cameraPreset: WorkDocumentEquationWordPresetCamera;
  lightRig: WorkDocumentEquationWordScene3DLightRig;
}

export type WorkDocumentEquationWordBevelPreset =
  | 'relaxedInset'
  | 'circle'
  | 'slope'
  | 'cross'
  | 'angle'
  | 'softRound'
  | 'convex'
  | 'coolSlant'
  | 'divot'
  | 'riblet'
  | 'hardEdge'
  | 'artDeco';

export type WorkDocumentEquationWordPresetMaterial =
  | 'legacyMatte'
  | 'legacyPlastic'
  | 'legacyMetal'
  | 'legacyWireframe'
  | 'matte'
  | 'plastic'
  | 'metal'
  | 'warmMatte'
  | 'translucentPowder'
  | 'powder'
  | 'darkEdge'
  | 'softEdge'
  | 'clear'
  | 'flat'
  | 'softMetal'
  | 'none';

export interface WorkDocumentEquationWordBevel {
  widthEmus?: number;
  heightEmus?: number;
  preset?: WorkDocumentEquationWordBevelPreset;
}

export interface WorkDocumentEquationWordProperties3D {
  extrusionHeightEmus?: number;
  contourWidthEmus?: number;
  materialPreset?: WorkDocumentEquationWordPresetMaterial;
  topBevel?: WorkDocumentEquationWordBevel;
  bottomBevel?: WorkDocumentEquationWordBevel;
  extrusionColor?: WorkDocumentEquationWordEffectColor;
  contourColor?: WorkDocumentEquationWordEffectColor;
}

export type WorkDocumentEquationWordLigatures =
  | 'none'
  | 'standard'
  | 'contextual'
  | 'historical'
  | 'discretional'
  | 'standardContextual'
  | 'standardHistorical'
  | 'contextualHistorical'
  | 'standardDiscretional'
  | 'contextualDiscretional'
  | 'historicalDiscretional'
  | 'standardContextualHistorical'
  | 'standardContextualDiscretional'
  | 'standardHistoricalDiscretional'
  | 'contextualHistoricalDiscretional'
  | 'all';

export type WorkDocumentEquationWordNumberForm =
  | 'default'
  | 'lining'
  | 'oldStyle';

export type WorkDocumentEquationWordNumberSpacing =
  | 'default'
  | 'proportional'
  | 'tabular';

export interface WorkDocumentEquationWordLineDash {
  preset?: WorkDocumentEquationWordPresetLineDash;
}

export type WorkDocumentEquationWordLineJoin =
  | { type: 'round' }
  | { type: 'bevel' }
  | { type: 'miter'; limitPercent?: number };

export interface WorkDocumentEquationWordTextOutlineEffect {
  widthEmus?: number;
  cap?: WorkDocumentEquationWordTextOutlineCap;
  compound?: WorkDocumentEquationWordTextOutlineCompound;
  alignment?: WorkDocumentEquationWordTextOutlineAlignment;
  fill?: WorkDocumentEquationWordEffectFill;
  dash?: WorkDocumentEquationWordLineDash;
  join?: WorkDocumentEquationWordLineJoin;
}

export interface WorkDocumentEquationWordUnderline {
  style: WorkDocumentEquationUnderlineStyle;
  color?: WorkDocumentEquationWordColor;
}

export interface WorkDocumentEquationWordRunBorder {
  style: WorkDocumentEquationWordLineBorderStyle;
  color?: WorkDocumentEquationWordColor;
  sizeEighthPoints?: number;
  spacingPoints?: number;
  shadow?: boolean;
  frame?: boolean;
}

export interface WorkDocumentEquationWordShading {
  pattern: WorkDocumentEquationWordShadingPattern;
  color?: WorkDocumentEquationWordColor;
  fill?: WorkDocumentEquationWordColor;
}

export interface WorkDocumentEquationWordFitText {
  widthTwips: number;
  id?: number;
}

export type WorkDocumentEquationWordVerticalAlignment =
  | 'baseline'
  | 'superscript'
  | 'subscript';

export type WorkDocumentEquationWordEmphasisMark =
  | 'none'
  | 'dot'
  | 'comma'
  | 'circle'
  | 'underDot';

export interface WorkDocumentEquationWordLanguages {
  latin?: string;
  eastAsia?: string;
  bidi?: string;
}

export type WorkDocumentEquationWordCombineBrackets =
  | 'none'
  | 'round'
  | 'square'
  | 'angle'
  | 'curly';

export interface WorkDocumentEquationWordEastAsianLayout {
  id?: number;
  combine?: boolean;
  combineBrackets?: WorkDocumentEquationWordCombineBrackets;
  vertical?: boolean;
  verticalCompress?: boolean;
}

export interface WorkDocumentEquationWordRunProperties {
  fonts?: WorkDocumentEquationWordRunFonts;
  bold?: boolean;
  boldComplexScript?: boolean;
  italic?: boolean;
  italicComplexScript?: boolean;
  allCaps?: boolean;
  smallCaps?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  noProof?: boolean;
  snapToGrid?: boolean;
  hidden?: boolean;
  webHidden?: boolean;
  color?: WorkDocumentEquationWordColor;
  characterSpacingTwips?: number;
  characterScalePercent?: number;
  kerningThresholdHalfPoints?: number;
  positionHalfPoints?: number;
  fontSize?: number;
  fontSizeComplexScript?: number;
  highlight?: WorkDocumentEquationWordHighlight;
  underline?: WorkDocumentEquationWordUnderline;
  textEffect?: WorkDocumentEquationWordTextEffect;
  border?: WorkDocumentEquationWordRunBorder;
  shading?: WorkDocumentEquationWordShading;
  fitText?: WorkDocumentEquationWordFitText;
  verticalAlignment?: WorkDocumentEquationWordVerticalAlignment;
  rightToLeft?: boolean;
  complexScript?: boolean;
  emphasisMark?: WorkDocumentEquationWordEmphasisMark;
  languages?: WorkDocumentEquationWordLanguages;
  eastAsianLayout?: WorkDocumentEquationWordEastAsianLayout;
  paragraphMarkAlwaysHidden?: boolean;
  glow?: WorkDocumentEquationWordGlow;
  shadowEffect?: WorkDocumentEquationWordShadowEffect;
  reflectionEffect?: WorkDocumentEquationWordReflectionEffect;
  textOutlineEffect?: WorkDocumentEquationWordTextOutlineEffect;
  textFillEffect?: WorkDocumentEquationWordTextFillEffect;
  scene3D?: WorkDocumentEquationWordScene3D;
  properties3D?: WorkDocumentEquationWordProperties3D;
  ligatures?: WorkDocumentEquationWordLigatures;
  numberForm?: WorkDocumentEquationWordNumberForm;
  numberSpacing?: WorkDocumentEquationWordNumberSpacing;
  stylisticSets?: number[];
}

export interface WorkDocumentEquationManualBreak {
  alignmentAt?: number;
}

interface WorkDocumentEquationControlRevisionIdentity {
  id: number;
  author: string;
  date?: string;
  dateUtc?: string;
}

type WorkDocumentEquationControlDeletionRevision =
  WorkDocumentEquationControlRevisionIdentity & {
    kind: 'deletion';
    child?: never;
  };

type WorkDocumentEquationControlInsertionRevision =
  WorkDocumentEquationControlRevisionIdentity & {
    kind: 'insertion';
    child?: WorkDocumentEquationControlDeletionRevision;
  };

export type WorkDocumentEquationControlRevision =
  | WorkDocumentEquationControlDeletionRevision
  | WorkDocumentEquationControlInsertionRevision
  | (WorkDocumentEquationControlRevisionIdentity & {
      kind: 'moveFrom' | 'moveTo';
      child?:
        | WorkDocumentEquationControlInsertionRevision
        | WorkDocumentEquationControlDeletionRevision;
    });

export interface WorkDocumentEquationArgumentProperties {
  size?: -2 | -1 | 1 | 2;
  controlProperties?: WorkDocumentEquationWordRunProperties;
  controlRevision?: WorkDocumentEquationControlRevision;
}

export type WorkDocumentEquationExpression =
  | {
      type: 'run';
      text: string;
      literal?: boolean;
      normalText?: boolean;
      script?: WorkDocumentEquationRunScript;
      style?: WorkDocumentEquationRunStyle;
      manualBreak?: WorkDocumentEquationManualBreak;
      alignment?: boolean;
      wordRunProperties?: WorkDocumentEquationWordRunProperties;
    }
  | {
      type: 'fraction';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      fractionType: WorkDocumentEquationFractionType;
      numerator: WorkDocumentEquationExpression[];
      numeratorProperties?: WorkDocumentEquationArgumentProperties;
      denominator: WorkDocumentEquationExpression[];
      denominatorProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'superscript';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      base: WorkDocumentEquationExpression[];
      baseProperties?: WorkDocumentEquationArgumentProperties;
      superScript: WorkDocumentEquationExpression[];
      superScriptProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'subscript';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      base: WorkDocumentEquationExpression[];
      baseProperties?: WorkDocumentEquationArgumentProperties;
      subScript: WorkDocumentEquationExpression[];
      subScriptProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'subSuperScript';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      alignScripts?: boolean;
      base: WorkDocumentEquationExpression[];
      baseProperties?: WorkDocumentEquationArgumentProperties;
      subScript: WorkDocumentEquationExpression[];
      subScriptProperties?: WorkDocumentEquationArgumentProperties;
      superScript: WorkDocumentEquationExpression[];
      superScriptProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'preSubSuperScript';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      base: WorkDocumentEquationExpression[];
      baseProperties?: WorkDocumentEquationArgumentProperties;
      subScript: WorkDocumentEquationExpression[];
      subScriptProperties?: WorkDocumentEquationArgumentProperties;
      superScript: WorkDocumentEquationExpression[];
      superScriptProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'lowerLimit';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      base: WorkDocumentEquationExpression[];
      baseProperties?: WorkDocumentEquationArgumentProperties;
      limit: WorkDocumentEquationExpression[];
      limitProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'upperLimit';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      base: WorkDocumentEquationExpression[];
      baseProperties?: WorkDocumentEquationArgumentProperties;
      limit: WorkDocumentEquationExpression[];
      limitProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'radical';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
      degree?: WorkDocumentEquationExpression[];
      degreeProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'function';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      name: WorkDocumentEquationExpression[];
      nameProperties?: WorkDocumentEquationArgumentProperties;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'nary';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      operator: WorkDocumentEquationNaryOperator;
      limitLocation: WorkDocumentEquationLimitLocation;
      grow?: boolean;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
      subScript?: WorkDocumentEquationExpression[];
      subScriptProperties?: WorkDocumentEquationArgumentProperties;
      superScript?: WorkDocumentEquationExpression[];
      superScriptProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'accent';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      character: string;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'bar';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      position: WorkDocumentEquationBarPosition;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'groupCharacter';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      character: string;
      position: WorkDocumentEquationBarPosition;
      verticalJustification: WorkDocumentEquationBarPosition;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'phantom';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      show: boolean;
      zeroWidth: boolean;
      zeroAscent: boolean;
      zeroDescent: boolean;
      transparent: boolean;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'borderBox';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      hideTop: boolean;
      hideBottom: boolean;
      hideLeft: boolean;
      hideRight: boolean;
      strikeHorizontal: boolean;
      strikeVertical: boolean;
      strikeBottomLeftToTopRight: boolean;
      strikeTopLeftToBottomRight: boolean;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'box';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      operatorEmulator: boolean;
      noBreak: boolean;
      differential: boolean;
      alignment: boolean;
      manualBreak?: WorkDocumentEquationManualBreak;
      children: WorkDocumentEquationExpression[];
      childrenProperties?: WorkDocumentEquationArgumentProperties;
    }
  | {
      type: 'matrix';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
      placeholdersHidden: boolean;
      columnAlignments: WorkDocumentEquationMatrixAlignment[];
      spacing?: WorkDocumentEquationMatrixSpacing;
      rows: WorkDocumentEquationExpression[][][];
      cellProperties?: Array<
        Array<WorkDocumentEquationArgumentProperties | null>
      >;
    }
  | {
      type: 'equationArray';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      baseAlignment: WorkDocumentEquationMatrixBaseAlignment;
      maximumDistribution: boolean;
      objectDistribution: boolean;
      rowSpacingRule: WorkDocumentEquationRowSpacingRule;
      rowSpacing: number;
      rows: WorkDocumentEquationExpression[][];
      rowProperties?: Array<WorkDocumentEquationArgumentProperties | null>;
    }
  | {
      type: 'delimiter';
      controlProperties?: WorkDocumentEquationWordRunProperties;
      controlRevision?: WorkDocumentEquationControlRevision;
      opening: string;
      closing: string;
      separator: string;
      grow?: boolean;
      shape?: WorkDocumentEquationDelimiterShape;
      arguments: WorkDocumentEquationExpression[][];
      argumentProperties?: Array<WorkDocumentEquationArgumentProperties | null>;
    };

export type WorkDocumentEquationNaryOperator =
  | '∑'
  | '∏'
  | '∐'
  | '∫'
  | '∬'
  | '∭'
  | '∮'
  | '∯'
  | '∰'
  | '⋂'
  | '⋃';

export interface WorkDocumentEquation {
  version: 1;
  display: WorkDocumentEquationDisplay;
  justification?: WorkDocumentEquationJustification;
  children: WorkDocumentEquationExpression[];
}

interface EquationNormalizationState {
  depth: number;
  nodes: number;
  textLength: number;
  equationArrayDepth: number;
  equationArrayAlignmentMarkers: number;
}

interface EquationArrayAlignmentState {
  markerIndex: number;
  started: boolean;
}

const EQUATION_SELECTOR = 'span[data-document-equation]';
const MATHML_NAMESPACE = 'http://www.w3.org/1998/Math/MathML';
const MAX_EQUATION_DEPTH = 32;
const MAX_EQUATION_NODES = 4_096;
const MAX_EQUATION_TEXT_LENGTH = 65_536;
const MAX_EQUATION_MODEL_LENGTH = 262_144;
const MAX_DELIMITER_ARGUMENTS = 32;
const MAX_MATRIX_ROWS = 64;
const MAX_MATRIX_COLUMNS = 64;
const MAX_MATRIX_CELLS = 1_024;
const MAX_MATRIX_SPACING = 65_535;
const MAX_MATRIX_MINIMUM_COLUMN_WIDTH = 31_680;
const MAX_EQUATION_ARRAY_ROWS = 64;
const MAX_EQUATION_ARRAY_ALIGNMENT_MARKERS = 4_096;
const NARY_OPERATORS = new Set<WorkDocumentEquationNaryOperator>([
  '∑',
  '∏',
  '∐',
  '∫',
  '∬',
  '∭',
  '∮',
  '∯',
  '∰',
  '⋂',
  '⋃',
]);
const FRACTION_TYPES = new Set<WorkDocumentEquationFractionType>([
  'bar',
  'noBar',
  'skewed',
  'linear',
]);
const RUN_SCRIPTS = new Set<WorkDocumentEquationRunScript>([
  'roman',
  'sansSerif',
  'monospace',
  'fraktur',
  'doubleStruck',
  'script',
]);
const RUN_STYLES = new Set<WorkDocumentEquationRunStyle>([
  'plain',
  'italic',
  'bold',
  'boldItalic',
]);
const THEME_FONTS = new Set<WorkDocumentEquationThemeFont>([
  'majorEastAsia',
  'majorBidi',
  'majorAscii',
  'majorHAnsi',
  'minorEastAsia',
  'minorBidi',
  'minorAscii',
  'minorHAnsi',
]);
const THEME_COLORS = new Set<WorkDocumentEquationThemeColor>([
  'dark1',
  'light1',
  'dark2',
  'light2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hyperlink',
  'followedHyperlink',
  'none',
  'background1',
  'text1',
  'background2',
  'text2',
]);
const WORD_EFFECT_SCHEME_COLORS =
  new Set<WorkDocumentEquationWordEffectSchemeColor>([
    'dark1',
    'light1',
    'dark2',
    'light2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hyperlink',
    'followedHyperlink',
    'background1',
    'text1',
    'background2',
    'text2',
    'placeholder',
  ]);
const WORD_COLOR_TRANSFORM_TYPES =
  new Set<WorkDocumentEquationWordColorTransformType>([
    'tint',
    'shade',
    'alpha',
    'hueMod',
    'saturation',
    'saturationOffset',
    'saturationModulation',
    'luminance',
    'luminanceOffset',
    'luminanceModulation',
  ]);
const WORD_FIXED_COLOR_TRANSFORM_TYPES =
  new Set<WorkDocumentEquationWordColorTransformType>([
    'tint',
    'shade',
    'alpha',
  ]);
const WORD_HIGHLIGHT_COLORS = new Set<WorkDocumentEquationWordHighlight>([
  'black',
  'blue',
  'cyan',
  'green',
  'magenta',
  'red',
  'yellow',
  'white',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
  'none',
]);
const WORD_HIGHLIGHT_MATHML_COLORS: Readonly<
  Record<Exclude<WorkDocumentEquationWordHighlight, 'none'>, string>
> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  yellow: '#ffff00',
  white: '#ffffff',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
};
const WORD_SHADING_PATTERNS = new Set<WorkDocumentEquationWordShadingPattern>([
  'nil',
  'clear',
  'solid',
  'horzStripe',
  'vertStripe',
  'reverseDiagStripe',
  'diagStripe',
  'horzCross',
  'diagCross',
  'thinHorzStripe',
  'thinVertStripe',
  'thinReverseDiagStripe',
  'thinDiagStripe',
  'thinHorzCross',
  'thinDiagCross',
  'pct5',
  'pct10',
  'pct12',
  'pct15',
  'pct20',
  'pct25',
  'pct30',
  'pct35',
  'pct37',
  'pct40',
  'pct45',
  'pct50',
  'pct55',
  'pct60',
  'pct62',
  'pct65',
  'pct70',
  'pct75',
  'pct80',
  'pct85',
  'pct87',
  'pct90',
  'pct95',
]);
const UNDERLINE_STYLES = new Set<WorkDocumentEquationUnderlineStyle>([
  'none',
  'words',
  'single',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dashDotHeavy',
  'dotDotDash',
  'dashDotDotHeavy',
  'wave',
  'wavyHeavy',
  'wavyDouble',
]);
const WORD_TEXT_EFFECTS = new Set<WorkDocumentEquationWordTextEffect>([
  'blinkBackground',
  'lights',
  'antsBlack',
  'antsRed',
  'shimmer',
  'sparkle',
  'none',
]);
const WORD_LINE_BORDER_STYLES =
  new Set<WorkDocumentEquationWordLineBorderStyle>([
    'nil',
    'none',
    'single',
    'thick',
    'double',
    'dotted',
    'dashed',
    'dotDash',
    'dotDotDash',
    'triple',
    'thinThickSmallGap',
    'thickThinSmallGap',
    'thinThickThinSmallGap',
    'thinThickMediumGap',
    'thickThinMediumGap',
    'thinThickThinMediumGap',
    'thinThickLargeGap',
    'thickThinLargeGap',
    'thinThickThinLargeGap',
    'wave',
    'doubleWave',
    'dashSmallGap',
    'dashDotStroked',
    'threeDEmboss',
    'threeDEngrave',
    'outset',
    'inset',
  ]);
const WORD_VERTICAL_ALIGNMENTS =
  new Set<WorkDocumentEquationWordVerticalAlignment>([
    'baseline',
    'superscript',
    'subscript',
  ]);
const WORD_EMPHASIS_MARKS = new Set<WorkDocumentEquationWordEmphasisMark>([
  'none',
  'dot',
  'comma',
  'circle',
  'underDot',
]);
const WORD_COMBINE_BRACKETS = new Set<WorkDocumentEquationWordCombineBrackets>([
  'none',
  'round',
  'square',
  'angle',
  'curly',
]);
const MAX_EQUATION_WORD_FONT_LENGTH = 127;
const MAX_EQUATION_LANGUAGE_LENGTH = 85;
const MAX_EQUATION_FONT_SIZE = 512;
const MAX_EQUATION_CHARACTER_SPACING_TWIPS = 31_680;
const MAX_EQUATION_CHARACTER_SCALE_PERCENT = 600;
const MAX_EQUATION_KERNING_THRESHOLD_HALF_POINTS = 3_277;
const MIN_EQUATION_WORD_LINE_BORDER_EIGHTH_POINTS = 2;
const MAX_EQUATION_WORD_LINE_BORDER_EIGHTH_POINTS = 96;
const MAX_EQUATION_WORD_BORDER_SPACING_POINTS = 31;
const MAX_EQUATION_WORD_FIT_TEXT_WIDTH_TWIPS = 31_680;
const MIN_EQUATION_WORD_FIT_TEXT_ID = -2_147_483_648;
const MAX_EQUATION_WORD_FIT_TEXT_ID = 2_147_483_647;
const MIN_EQUATION_WORD_EAST_ASIAN_LAYOUT_ID = -2_147_483_648;
const MAX_EQUATION_WORD_EAST_ASIAN_LAYOUT_ID = 2_147_483_647;
const MAX_EQUATION_WORD_GLOW_RADIUS_EMUS = 2_147_483_647;
const MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS = 2_147_483_647;
const MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS = 21_599_999;
const MIN_EQUATION_WORD_EFFECT_SCALE_UNITS = -2_147_483_648;
const MAX_EQUATION_WORD_EFFECT_SCALE_UNITS = 2_147_483_647;
const MIN_EQUATION_WORD_EFFECT_SKEW_UNITS = -5_399_999;
const MAX_EQUATION_WORD_EFFECT_SKEW_UNITS = 5_399_999;
const MAX_EQUATION_WORD_FIXED_PERCENTAGE_UNITS = 100_000;
const MAX_EQUATION_WORD_TEXT_OUTLINE_WIDTH_EMUS = 20_116_800;
const MIN_EQUATION_WORD_GRADIENT_STOPS = 2;
const MAX_EQUATION_WORD_GRADIENT_STOPS = 10;
const EQUATION_WORD_ANGLE_UNITS_PER_DEGREE = 60_000;
const EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT = 1_000;
const MAX_EQUATION_WORD_COLOR_TRANSFORMS = 64;
const MAX_EQUATION_WORD_STYLISTIC_SET_ENTRIES = 4_096;
const MIN_EQUATION_WORD_STYLISTIC_SET_ID = 1;
const MAX_EQUATION_WORD_STYLISTIC_SET_ID = 20;
const MIN_EQUATION_WORD_COLOR_PERCENTAGE = -2_147_483_648;
const MAX_EQUATION_WORD_COLOR_PERCENTAGE = 2_147_483_647;
const MAX_EQUATION_WORD_FIXED_COLOR_PERCENTAGE = 100_000;
const MIN_EQUATION_POSITION_HALF_POINTS = -2_147_483_648;
const MAX_EQUATION_POSITION_HALF_POINTS = 2_147_483_647;
const MAX_EQUATION_CONTROL_REVISION_ID = 2_147_483_647;
const MAX_EQUATION_CONTROL_REVISION_AUTHOR_LENGTH = 255;
const MAX_EQUATION_CONTROL_REVISION_DATE_LENGTH = 64;
const CONTROL_REVISION_KEYS = new Set([
  'kind',
  'id',
  'author',
  'date',
  'dateUtc',
  'child',
]);
const CONTROL_REVISION_KINDS = new Set<
  WorkDocumentEquationControlRevision['kind']
>(['insertion', 'deletion', 'moveFrom', 'moveTo']);
const CONTROL_REVISION_INSERTION_CHILD_KINDS = new Set<
  WorkDocumentEquationControlRevision['kind']
>(['deletion']);
const CONTROL_REVISION_MOVE_CHILD_KINDS = new Set<
  WorkDocumentEquationControlRevision['kind']
>(['insertion', 'deletion']);
const WORD_RUN_PROPERTY_KEYS = new Set([
  'fonts',
  'bold',
  'boldComplexScript',
  'italic',
  'italicComplexScript',
  'allCaps',
  'smallCaps',
  'strike',
  'doubleStrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'hidden',
  'webHidden',
  'color',
  'characterSpacingTwips',
  'characterScalePercent',
  'kerningThresholdHalfPoints',
  'positionHalfPoints',
  'fontSize',
  'fontSizeComplexScript',
  'highlight',
  'underline',
  'textEffect',
  'border',
  'shading',
  'fitText',
  'verticalAlignment',
  'rightToLeft',
  'complexScript',
  'emphasisMark',
  'languages',
  'eastAsianLayout',
  'paragraphMarkAlwaysHidden',
  'glow',
  'shadowEffect',
  'reflectionEffect',
  'textOutlineEffect',
  'textFillEffect',
  'scene3D',
  'properties3D',
  'ligatures',
  'numberForm',
  'numberSpacing',
  'stylisticSets',
]);
const WORD_RUN_FONT_KEYS = new Set([
  'ascii',
  'highAnsi',
  'eastAsia',
  'complexScript',
  'asciiTheme',
  'highAnsiTheme',
  'eastAsiaTheme',
  'complexScriptTheme',
  'hint',
]);
const WORD_COLOR_KEYS = new Set(['value', 'theme', 'tint', 'shade']);
const WORD_UNDERLINE_KEYS = new Set(['style', 'color']);
const WORD_RUN_BORDER_KEYS = new Set([
  'style',
  'color',
  'sizeEighthPoints',
  'spacingPoints',
  'shadow',
  'frame',
]);
const WORD_SHADING_KEYS = new Set(['pattern', 'color', 'fill']);
const WORD_FIT_TEXT_KEYS = new Set(['widthTwips', 'id']);
const WORD_LANGUAGE_KEYS = new Set(['latin', 'eastAsia', 'bidi']);
const WORD_EAST_ASIAN_LAYOUT_KEYS = new Set([
  'id',
  'combine',
  'combineBrackets',
  'vertical',
  'verticalCompress',
]);
const WORD_GLOW_KEYS = new Set(['radiusEmus', 'color']);
const WORD_SHADOW_EFFECT_KEYS = new Set([
  'blurRadiusEmus',
  'distanceEmus',
  'directionDegrees',
  'horizontalScalePercent',
  'verticalScalePercent',
  'horizontalSkewDegrees',
  'verticalSkewDegrees',
  'alignment',
  'color',
]);
const WORD_REFLECTION_EFFECT_KEYS = new Set([
  'blurRadiusEmus',
  'startOpacityPercent',
  'startPositionPercent',
  'endOpacityPercent',
  'endPositionPercent',
  'distanceEmus',
  'directionDegrees',
  'fadeDirectionDegrees',
  'horizontalScalePercent',
  'verticalScalePercent',
  'horizontalSkewDegrees',
  'verticalSkewDegrees',
  'alignment',
]);
const WORD_TEXT_OUTLINE_EFFECT_KEYS = new Set([
  'widthEmus',
  'cap',
  'compound',
  'alignment',
  'fill',
  'dash',
  'join',
]);
const WORD_TEXT_FILL_EFFECT_KEYS = new Set(['fill']);
const WORD_SCENE_3D_KEYS = new Set(['cameraPreset', 'lightRig']);
const WORD_SCENE_3D_LIGHT_RIG_KEYS = new Set([
  'preset',
  'direction',
  'rotation',
]);
const WORD_SCENE_3D_ROTATION_KEYS = new Set([
  'latitudeDegrees',
  'longitudeDegrees',
  'revolutionDegrees',
]);
const WORD_PROPERTIES_3D_KEYS = new Set([
  'extrusionHeightEmus',
  'contourWidthEmus',
  'materialPreset',
  'topBevel',
  'bottomBevel',
  'extrusionColor',
  'contourColor',
]);
const WORD_PROPERTIES_3D_BEVEL_KEYS = new Set([
  'widthEmus',
  'heightEmus',
  'preset',
]);
const WORD_PROPERTIES_3D_BEVEL_PRESETS =
  new Set<WorkDocumentEquationWordBevelPreset>([
    'relaxedInset',
    'circle',
    'slope',
    'cross',
    'angle',
    'softRound',
    'convex',
    'coolSlant',
    'divot',
    'riblet',
    'hardEdge',
    'artDeco',
  ]);
const WORD_PROPERTIES_3D_MATERIAL_PRESETS =
  new Set<WorkDocumentEquationWordPresetMaterial>([
    'legacyMatte',
    'legacyPlastic',
    'legacyMetal',
    'legacyWireframe',
    'matte',
    'plastic',
    'metal',
    'warmMatte',
    'translucentPowder',
    'powder',
    'darkEdge',
    'softEdge',
    'clear',
    'flat',
    'softMetal',
    'none',
  ]);
const WORD_LIGATURE_STANDARD = 1;
const WORD_LIGATURE_CONTEXTUAL = 2;
const WORD_LIGATURE_HISTORICAL = 4;
const WORD_LIGATURE_DISCRETIONAL = 8;
const WORD_LIGATURE_FLAGS = new Map<WorkDocumentEquationWordLigatures, number>([
  ['none', 0],
  ['standard', WORD_LIGATURE_STANDARD],
  ['contextual', WORD_LIGATURE_CONTEXTUAL],
  ['historical', WORD_LIGATURE_HISTORICAL],
  ['discretional', WORD_LIGATURE_DISCRETIONAL],
  ['standardContextual', WORD_LIGATURE_STANDARD | WORD_LIGATURE_CONTEXTUAL],
  ['standardHistorical', WORD_LIGATURE_STANDARD | WORD_LIGATURE_HISTORICAL],
  ['contextualHistorical', WORD_LIGATURE_CONTEXTUAL | WORD_LIGATURE_HISTORICAL],
  ['standardDiscretional', WORD_LIGATURE_STANDARD | WORD_LIGATURE_DISCRETIONAL],
  [
    'contextualDiscretional',
    WORD_LIGATURE_CONTEXTUAL | WORD_LIGATURE_DISCRETIONAL,
  ],
  [
    'historicalDiscretional',
    WORD_LIGATURE_HISTORICAL | WORD_LIGATURE_DISCRETIONAL,
  ],
  [
    'standardContextualHistorical',
    WORD_LIGATURE_STANDARD |
      WORD_LIGATURE_CONTEXTUAL |
      WORD_LIGATURE_HISTORICAL,
  ],
  [
    'standardContextualDiscretional',
    WORD_LIGATURE_STANDARD |
      WORD_LIGATURE_CONTEXTUAL |
      WORD_LIGATURE_DISCRETIONAL,
  ],
  [
    'standardHistoricalDiscretional',
    WORD_LIGATURE_STANDARD |
      WORD_LIGATURE_HISTORICAL |
      WORD_LIGATURE_DISCRETIONAL,
  ],
  [
    'contextualHistoricalDiscretional',
    WORD_LIGATURE_CONTEXTUAL |
      WORD_LIGATURE_HISTORICAL |
      WORD_LIGATURE_DISCRETIONAL,
  ],
  [
    'all',
    WORD_LIGATURE_STANDARD |
      WORD_LIGATURE_CONTEXTUAL |
      WORD_LIGATURE_HISTORICAL |
      WORD_LIGATURE_DISCRETIONAL,
  ],
]);
const WORD_NUMBER_FORMS = new Set<WorkDocumentEquationWordNumberForm>([
  'default',
  'lining',
  'oldStyle',
]);
const WORD_NUMBER_SPACINGS = new Set<WorkDocumentEquationWordNumberSpacing>([
  'default',
  'proportional',
  'tabular',
]);
const WORD_SCENE_3D_CAMERA_PRESETS =
  new Set<WorkDocumentEquationWordPresetCamera>([
    'legacyObliqueTopLeft',
    'legacyObliqueTop',
    'legacyObliqueTopRight',
    'legacyObliqueLeft',
    'legacyObliqueFront',
    'legacyObliqueRight',
    'legacyObliqueBottomLeft',
    'legacyObliqueBottom',
    'legacyObliqueBottomRight',
    'legacyPerspectiveTopLeft',
    'legacyPerspectiveTop',
    'legacyPerspectiveTopRight',
    'legacyPerspectiveLeft',
    'legacyPerspectiveFront',
    'legacyPerspectiveRight',
    'legacyPerspectiveBottomLeft',
    'legacyPerspectiveBottom',
    'legacyPerspectiveBottomRight',
    'orthographicFront',
    'isometricTopUp',
    'isometricTopDown',
    'isometricBottomUp',
    'isometricBottomDown',
    'isometricLeftUp',
    'isometricLeftDown',
    'isometricRightUp',
    'isometricRightDown',
    'isometricOffAxis1Left',
    'isometricOffAxis1Right',
    'isometricOffAxis1Top',
    'isometricOffAxis2Left',
    'isometricOffAxis2Right',
    'isometricOffAxis2Top',
    'isometricOffAxis3Left',
    'isometricOffAxis3Right',
    'isometricOffAxis3Bottom',
    'isometricOffAxis4Left',
    'isometricOffAxis4Right',
    'isometricOffAxis4Bottom',
    'obliqueTopLeft',
    'obliqueTop',
    'obliqueTopRight',
    'obliqueLeft',
    'obliqueRight',
    'obliqueBottomLeft',
    'obliqueBottom',
    'obliqueBottomRight',
    'perspectiveFront',
    'perspectiveLeft',
    'perspectiveRight',
    'perspectiveAbove',
    'perspectiveBelow',
    'perspectiveAboveLeftFacing',
    'perspectiveAboveRightFacing',
    'perspectiveContrastingLeftFacing',
    'perspectiveContrastingRightFacing',
    'perspectiveHeroicLeftFacing',
    'perspectiveHeroicRightFacing',
    'perspectiveHeroicExtremeLeftFacing',
    'perspectiveHeroicExtremeRightFacing',
    'perspectiveRelaxed',
    'perspectiveRelaxedModerately',
  ]);
const WORD_SCENE_3D_LIGHT_RIG_PRESETS =
  new Set<WorkDocumentEquationWordLightRigPreset>([
    'legacyFlat1',
    'legacyFlat2',
    'legacyFlat3',
    'legacyFlat4',
    'legacyNormal1',
    'legacyNormal2',
    'legacyNormal3',
    'legacyNormal4',
    'legacyHarsh1',
    'legacyHarsh2',
    'legacyHarsh3',
    'legacyHarsh4',
    'threePoint',
    'balanced',
    'soft',
    'harsh',
    'flood',
    'contrasting',
    'morning',
    'sunrise',
    'sunset',
    'chilly',
    'freezing',
    'flat',
    'twoPoint',
    'glow',
    'brightRoom',
  ]);
const WORD_SCENE_3D_LIGHT_RIG_DIRECTIONS =
  new Set<WorkDocumentEquationWordLightRigDirection>([
    'topLeft',
    'top',
    'topRight',
    'left',
    'right',
    'bottomLeft',
    'bottom',
    'bottomRight',
  ]);
const WORD_EFFECT_NO_FILL_KEYS = new Set(['type']);
const WORD_EFFECT_SOLID_FILL_KEYS = new Set(['type', 'color']);
const WORD_EFFECT_GRADIENT_FILL_KEYS = new Set(['type', 'stops', 'shade']);
const WORD_GRADIENT_STOP_KEYS = new Set(['positionPercent', 'color']);
const WORD_LINEAR_GRADIENT_SHADE_KEYS = new Set([
  'type',
  'angleDegrees',
  'scaled',
]);
const WORD_PATH_GRADIENT_SHADE_KEYS = new Set([
  'type',
  'path',
  'fillToRectangle',
]);
const WORD_GRADIENT_FILL_RECTANGLE_KEYS = new Set([
  'leftPercent',
  'topPercent',
  'rightPercent',
  'bottomPercent',
]);
const WORD_LINE_DASH_KEYS = new Set(['preset']);
const WORD_ROUND_OR_BEVEL_LINE_JOIN_KEYS = new Set(['type']);
const WORD_MITER_LINE_JOIN_KEYS = new Set(['type', 'limitPercent']);
const WORD_RECTANGLE_ALIGNMENTS =
  new Set<WorkDocumentEquationWordRectangleAlignment>([
    'none',
    'topLeft',
    'top',
    'topRight',
    'left',
    'center',
    'right',
    'bottomLeft',
    'bottom',
    'bottomRight',
  ]);
const WORD_TEXT_OUTLINE_CAPS = new Set<WorkDocumentEquationWordTextOutlineCap>([
  'round',
  'square',
  'flat',
]);
const WORD_TEXT_OUTLINE_COMPOUNDS =
  new Set<WorkDocumentEquationWordTextOutlineCompound>([
    'single',
    'double',
    'thickThin',
    'thinThick',
    'triple',
  ]);
const WORD_TEXT_OUTLINE_ALIGNMENTS =
  new Set<WorkDocumentEquationWordTextOutlineAlignment>(['center', 'inset']);
const WORD_PRESET_LINE_DASHES = new Set<WorkDocumentEquationWordPresetLineDash>(
  [
    'solid',
    'dot',
    'systemDot',
    'dash',
    'systemDash',
    'longDash',
    'dashDot',
    'systemDashDot',
    'longDashDot',
    'longDashDotDot',
    'systemDashDotDot',
  ],
);
const WORD_GRADIENT_PATHS = new Set<WorkDocumentEquationWordGradientPath>([
  'shape',
  'circle',
  'rectangle',
]);
const WORD_EFFECT_COLOR_KEYS = new Set(['type', 'value', 'transforms']);
const WORD_COLOR_TRANSFORM_KEYS = new Set(['type', 'value']);
const LIMIT_LOCATIONS = new Set<WorkDocumentEquationLimitLocation>([
  'underOver',
  'subSup',
]);
const DELIMITER_SHAPES = new Set<WorkDocumentEquationDelimiterShape>([
  'centered',
  'match',
]);
const EQUATION_JUSTIFICATIONS = new Set<WorkDocumentEquationJustification>([
  'left',
  'right',
  'center',
  'centerGroup',
]);
const MATRIX_ALIGNMENTS = new Set<WorkDocumentEquationMatrixAlignment>([
  'left',
  'center',
  'right',
]);
const MATRIX_BASE_ALIGNMENTS = new Set<WorkDocumentEquationMatrixBaseAlignment>(
  ['top', 'center', 'bottom'],
);
const EQUATION_SPACING_RULES = new Set<WorkDocumentEquationRowSpacingRule>([
  'single',
  'oneAndHalf',
  'double',
  'exact',
  'multiple',
]);
const MATRIX_SPACING_KEYS = new Set([
  'rowSpacingRule',
  'rowSpacing',
  'columnGapRule',
  'columnGap',
  'minimumColumnWidthTwips',
]);
const BAR_POSITIONS = new Set<WorkDocumentEquationBarPosition>([
  'top',
  'bottom',
]);
const MATHML_ACCENT_CHARACTERS = new Map([
  ['\u0300', '`'],
  ['\u0301', '\u00b4'],
  ['\u0302', '\u02c6'],
  ['\u0303', '\u02dc'],
  ['\u0304', '\u00af'],
  ['\u0305', '\u203e'],
  ['\u0306', '\u02d8'],
  ['\u0307', '\u02d9'],
  ['\u0308', '\u00a8'],
  ['\u030a', '\u02da'],
  ['\u030b', '\u02dd'],
  ['\u030c', '\u02c7'],
  ['\u20d6', '\u2190'],
  ['\u20d7', '\u2192'],
  ['\u20e1', '\u2194'],
]);

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentEquation: {
      insertDocumentEquation: (equation: WorkDocumentEquation) => ReturnType;
      updateDocumentEquation: (equation: WorkDocumentEquation) => ReturnType;
    };
  }
}

export const DocumentEquation = Node.create({
  name: 'documentEquation',
  inline: true,
  group: 'inline',
  atom: true,
  selectable: true,

  addCommands() {
    return {
      insertDocumentEquation: (equation) => (props) =>
        insertDocumentEquationCommand(props, equation),
      updateDocumentEquation: (equation) => (props) =>
        updateDocumentEquationCommand(props, equation),
    };
  },

  addAttributes() {
    return {
      equation: {
        default: defaultDocumentEquation(),
        rendered: false,
      },
    };
  },

  parseHTML() {
    return [
      {
        tag: EQUATION_SELECTOR,
        priority: 120,
        getAttrs: (node) => {
          if (!(node instanceof HTMLElement)) return false;
          const equation = documentEquationFromElement(node);
          return equation ? { equation } : false;
        },
      },
    ];
  },

  renderHTML({ node, HTMLAttributes }) {
    const equation =
      normalizeDocumentEquation(node.attrs.equation) ??
      defaultDocumentEquation();
    const label = documentEquationText(equation);
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-equation': 'true',
        'data-equation-display': equation.display,
        'data-equation-model': serializeDocumentEquation(equation),
        'aria-label': label,
        class: documentEquationClassName(equation),
        contenteditable: 'false',
        role: 'math',
      }),
      equationMathMl(equation),
    ];
  },

  renderText({ node }) {
    return documentEquationText(
      normalizeDocumentEquation(node.attrs.equation) ??
        defaultDocumentEquation(),
    );
  },
});

export function normalizeDocumentEquation(
  source: unknown,
): WorkDocumentEquation | null {
  if (!isRecord(source) || source.version !== 1) return null;
  const display =
    source.display === 'block' || source.display === 'inline'
      ? source.display
      : null;
  if (!display) return null;
  const justification =
    source.justification === undefined
      ? undefined
      : EQUATION_JUSTIFICATIONS.has(
            source.justification as WorkDocumentEquationJustification,
          )
        ? (source.justification as WorkDocumentEquationJustification)
        : null;
  if (
    justification === null ||
    (display === 'inline' && source.justification !== undefined)
  ) {
    return null;
  }
  const state: EquationNormalizationState = {
    depth: 0,
    nodes: 0,
    textLength: 0,
    equationArrayDepth: 0,
    equationArrayAlignmentMarkers: 0,
  };
  const children = normalizeExpressionList(source.children, state);
  if (!children) return null;
  const equation = {
    version: 1,
    display,
    ...(display === 'block' && justification && justification !== 'centerGroup'
      ? { justification }
      : {}),
    children,
  } satisfies WorkDocumentEquation;
  return JSON.stringify(equation).length <= MAX_EQUATION_MODEL_LENGTH
    ? equation
    : null;
}

export function serializeDocumentEquation(
  source: WorkDocumentEquation,
): string {
  const equation = normalizeDocumentEquation(source);
  if (!equation) throw new Error('The document equation model is invalid.');
  const serialized = JSON.stringify(equation);
  if (serialized.length > MAX_EQUATION_MODEL_LENGTH) {
    throw new Error('The document equation model exceeds the size limit.');
  }
  return serialized;
}

export function parseDocumentEquation(
  source: string | undefined,
): WorkDocumentEquation | null {
  if (!source || source.length > MAX_EQUATION_MODEL_LENGTH) return null;
  try {
    return normalizeDocumentEquation(JSON.parse(source));
  } catch {
    return null;
  }
}

export function documentEquationFromElement(
  element: Element,
): WorkDocumentEquation | null {
  if (!(element instanceof HTMLElement)) return null;
  return parseDocumentEquation(element.dataset.equationModel);
}

export function createDocumentEquationElement(
  document: Document,
  source: WorkDocumentEquation,
): HTMLElement {
  const equation = normalizeDocumentEquation(source);
  if (!equation) throw new Error('The document equation model is invalid.');
  const element = document.createElement('span');
  const label = documentEquationText(equation);
  element.dataset.documentEquation = 'true';
  element.dataset.equationDisplay = equation.display;
  element.dataset.equationModel = serializeDocumentEquation(equation);
  element.className = documentEquationClassName(equation);
  element.contentEditable = 'false';
  element.setAttribute('role', 'math');
  element.setAttribute('aria-label', label);
  element.append(
    createEquationDom(element.ownerDocument, equationMathMl(equation)),
  );
  return element;
}

export function documentEquationText(source: WorkDocumentEquation): string {
  return expressionListText(source.children).trim() || 'Equation';
}

export function defaultDocumentEquation(): WorkDocumentEquation {
  return {
    version: 1,
    display: 'inline',
    children: [{ type: 'run', text: 'x' }],
  };
}

function documentEquationClassName(equation: WorkDocumentEquation): string {
  if (equation.display === 'inline') return 'work-document-equation inline';
  const justification = equation.justification ?? 'centerGroup';
  const className =
    justification === 'centerGroup' ? 'center-group' : justification;
  return `work-document-equation block justification-${className}`;
}

function insertDocumentEquationCommand(
  { dispatch, editor, tr }: CommandProps,
  source: WorkDocumentEquation,
): boolean {
  const equation = normalizeDocumentEquation(source);
  const equationType = editor.schema.nodes.documentEquation;
  if (!equation || !equationType) return false;
  if (!dispatch) return true;
  tr.replaceSelectionWith(equationType.create({ equation }), false);
  tr.scrollIntoView();
  return true;
}

function updateDocumentEquationCommand(
  { dispatch, editor, state, tr }: CommandProps,
  source: WorkDocumentEquation,
): boolean {
  const equation = normalizeDocumentEquation(source);
  const equationType = editor.schema.nodes.documentEquation;
  const position = state.selection.from;
  const selected = state.doc.nodeAt(position);
  if (!equation || !equationType || selected?.type !== equationType) {
    return false;
  }
  if (!dispatch) return true;
  tr.setNodeMarkup(position, equationType, { equation });
  tr.scrollIntoView();
  return true;
}

function normalizeExpressionList(
  source: unknown,
  state: EquationNormalizationState,
  optional = false,
): WorkDocumentEquationExpression[] | null {
  if (!Array.isArray(source) || source.length > MAX_EQUATION_NODES) return null;
  if (!optional && source.length === 0) return null;
  const children: WorkDocumentEquationExpression[] = [];
  for (const child of source) {
    const normalized = normalizeExpression(child, state);
    if (!normalized) return null;
    children.push(normalized);
  }
  return children;
}

function normalizeMathArgument(
  source: unknown,
  state: EquationNormalizationState,
): WorkDocumentEquationExpression[] | null {
  return normalizeExpressionList(source, state, true);
}

function normalizeMathArgumentProperties(
  source: unknown,
): WorkDocumentEquationArgumentProperties | null | undefined {
  if (source === undefined) return undefined;
  if (
    !isRecord(source) ||
    Object.keys(source).some(
      (key) =>
        key !== 'size' &&
        key !== 'controlProperties' &&
        key !== 'controlRevision',
    )
  ) {
    return null;
  }
  if (
    source.size !== undefined &&
    (typeof source.size !== 'number' ||
      !Number.isInteger(source.size) ||
      source.size < -2 ||
      source.size > 2)
  ) {
    return null;
  }
  const size = source.size === 0 ? undefined : (source.size as -2 | -1 | 1 | 2);
  const controlProperties =
    source.controlProperties === undefined
      ? undefined
      : normalizeEquationWordRunProperties(source.controlProperties);
  const controlRevision =
    source.controlRevision === undefined
      ? undefined
      : normalizeEquationControlRevision(source.controlRevision);
  if (controlProperties === null || controlRevision === null) return null;
  return size !== undefined || controlProperties || controlRevision
    ? {
        ...(size !== undefined ? { size } : {}),
        ...(controlProperties ? { controlProperties } : {}),
        ...(controlRevision ? { controlRevision } : {}),
      }
    : undefined;
}

function normalizeMathArgumentPropertySlots(
  source: unknown,
  length: number,
): Array<WorkDocumentEquationArgumentProperties | null> | null | undefined {
  if (source === undefined) return undefined;
  if (!Array.isArray(source) || source.length !== length) return null;
  const properties: Array<WorkDocumentEquationArgumentProperties | null> = [];
  let present = false;
  for (const candidate of source) {
    if (candidate === null) {
      properties.push(null);
      continue;
    }
    const normalized = normalizeMathArgumentProperties(candidate);
    if (normalized === null) return null;
    properties.push(normalized ?? null);
    present ||= Boolean(normalized);
  }
  return present ? properties : undefined;
}

function normalizeMatrixArgumentPropertySlots(
  source: unknown,
  rows: readonly (readonly unknown[])[],
):
  | Array<Array<WorkDocumentEquationArgumentProperties | null>>
  | null
  | undefined {
  if (source === undefined) return undefined;
  if (!Array.isArray(source) || source.length !== rows.length) return null;
  const properties: Array<
    Array<WorkDocumentEquationArgumentProperties | null>
  > = [];
  let present = false;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = source[rowIndex];
    if (!Array.isArray(row) || row.length !== rows[rowIndex]?.length) {
      return null;
    }
    const normalizedRow: Array<WorkDocumentEquationArgumentProperties | null> =
      [];
    for (const candidate of row) {
      if (candidate === null) {
        normalizedRow.push(null);
        continue;
      }
      const normalized = normalizeMathArgumentProperties(candidate);
      if (normalized === null) return null;
      normalizedRow.push(normalized ?? null);
      present ||= Boolean(normalized);
    }
    properties.push(normalizedRow);
  }
  return present ? properties : undefined;
}

function normalizeMatrixSpacing(
  source: unknown,
): WorkDocumentEquationMatrixSpacing | null {
  if (!isRecordWithKeys(source, MATRIX_SPACING_KEYS)) return null;
  const rowSpacingRule = EQUATION_SPACING_RULES.has(
    source.rowSpacingRule as WorkDocumentEquationSpacingRule,
  )
    ? (source.rowSpacingRule as WorkDocumentEquationSpacingRule)
    : null;
  const columnGapRule = EQUATION_SPACING_RULES.has(
    source.columnGapRule as WorkDocumentEquationSpacingRule,
  )
    ? (source.columnGapRule as WorkDocumentEquationSpacingRule)
    : null;
  const rowSpacing = source.rowSpacing;
  const columnGap = source.columnGap;
  const minimumColumnWidthTwips = source.minimumColumnWidthTwips;
  if (
    !rowSpacingRule ||
    !columnGapRule ||
    !boundedInteger(rowSpacing, 0, MAX_MATRIX_SPACING) ||
    !boundedInteger(columnGap, 0, MAX_MATRIX_SPACING) ||
    !boundedInteger(minimumColumnWidthTwips, 0, MAX_MATRIX_MINIMUM_COLUMN_WIDTH)
  ) {
    return null;
  }
  return {
    rowSpacingRule,
    rowSpacing,
    columnGapRule,
    columnGap,
    minimumColumnWidthTwips,
  };
}

function normalizeExpression(
  source: unknown,
  state: EquationNormalizationState,
): WorkDocumentEquationExpression | null {
  if (
    !isRecord(source) ||
    state.depth >= MAX_EQUATION_DEPTH ||
    state.nodes >= MAX_EQUATION_NODES
  ) {
    return null;
  }
  state.nodes += 1;
  state.depth += 1;
  try {
    if (source.type === 'run') {
      if (
        source.controlProperties !== undefined ||
        source.controlRevision !== undefined
      ) {
        return null;
      }
      if (
        typeof source.text !== 'string' ||
        source.text.length === 0 ||
        !validXmlText(source.text)
      )
        return null;
      state.textLength += source.text.length;
      if (state.textLength > MAX_EQUATION_TEXT_LENGTH) return null;
      if (state.equationArrayDepth > 0) {
        state.equationArrayAlignmentMarkers +=
          source.text.split('&').length - 1;
        if (
          state.equationArrayAlignmentMarkers >
          MAX_EQUATION_ARRAY_ALIGNMENT_MARKERS
        ) {
          return null;
        }
      }
      const literal = source.literal === undefined ? false : source.literal;
      const normalText =
        source.normalText === undefined ? false : source.normalText;
      const alignment =
        source.alignment === undefined ? false : source.alignment;
      const script =
        source.script === undefined
          ? undefined
          : RUN_SCRIPTS.has(source.script as WorkDocumentEquationRunScript)
            ? (source.script as WorkDocumentEquationRunScript)
            : null;
      const style =
        source.style === undefined
          ? undefined
          : RUN_STYLES.has(source.style as WorkDocumentEquationRunStyle)
            ? (source.style as WorkDocumentEquationRunStyle)
            : null;
      const manualBreak =
        source.manualBreak === undefined
          ? undefined
          : normalizeManualBreak(source.manualBreak);
      const wordRunProperties =
        source.wordRunProperties === undefined
          ? undefined
          : normalizeEquationWordRunProperties(source.wordRunProperties);
      if (
        typeof literal !== 'boolean' ||
        typeof normalText !== 'boolean' ||
        typeof alignment !== 'boolean' ||
        script === null ||
        style === null ||
        manualBreak === null ||
        wordRunProperties === null
      ) {
        return null;
      }
      return {
        type: 'run',
        text: source.text,
        ...(literal ? { literal } : {}),
        ...(normalText ? { normalText } : {}),
        ...(script && script !== 'roman' ? { script } : {}),
        ...(style && style !== 'italic' ? { style } : {}),
        ...(manualBreak ? { manualBreak } : {}),
        ...(alignment ? { alignment } : {}),
        ...(wordRunProperties ? { wordRunProperties } : {}),
      };
    }
    const controlProperties =
      source.controlProperties === undefined
        ? undefined
        : normalizeEquationWordRunProperties(source.controlProperties);
    const controlRevision =
      source.controlRevision === undefined
        ? undefined
        : normalizeEquationControlRevision(source.controlRevision);
    if (controlProperties === null || controlRevision === null) return null;
    const normalizedControlProperties = {
      ...(controlProperties ? { controlProperties } : {}),
      ...(controlRevision ? { controlRevision } : {}),
    };
    if (source.type === 'fraction') {
      const fractionType = FRACTION_TYPES.has(
        source.fractionType as WorkDocumentEquationFractionType,
      )
        ? (source.fractionType as WorkDocumentEquationFractionType)
        : null;
      const numerator = normalizeMathArgument(source.numerator, state);
      const denominator = normalizeMathArgument(source.denominator, state);
      const numeratorProperties = normalizeMathArgumentProperties(
        source.numeratorProperties,
      );
      const denominatorProperties = normalizeMathArgumentProperties(
        source.denominatorProperties,
      );
      return fractionType &&
        numerator &&
        denominator &&
        numeratorProperties !== null &&
        denominatorProperties !== null
        ? {
            type: 'fraction',
            ...normalizedControlProperties,
            fractionType,
            numerator,
            ...(numeratorProperties ? { numeratorProperties } : {}),
            denominator,
            ...(denominatorProperties ? { denominatorProperties } : {}),
          }
        : null;
    }
    if (source.type === 'superscript') {
      const base = normalizeMathArgument(source.base, state);
      const superScript = normalizeMathArgument(source.superScript, state);
      const baseProperties = normalizeMathArgumentProperties(
        source.baseProperties,
      );
      const superScriptProperties = normalizeMathArgumentProperties(
        source.superScriptProperties,
      );
      return base &&
        superScript &&
        baseProperties !== null &&
        superScriptProperties !== null
        ? {
            type: 'superscript',
            ...normalizedControlProperties,
            base,
            ...(baseProperties ? { baseProperties } : {}),
            superScript,
            ...(superScriptProperties ? { superScriptProperties } : {}),
          }
        : null;
    }
    if (source.type === 'subscript') {
      const base = normalizeMathArgument(source.base, state);
      const subScript = normalizeMathArgument(source.subScript, state);
      const baseProperties = normalizeMathArgumentProperties(
        source.baseProperties,
      );
      const subScriptProperties = normalizeMathArgumentProperties(
        source.subScriptProperties,
      );
      return base &&
        subScript &&
        baseProperties !== null &&
        subScriptProperties !== null
        ? {
            type: 'subscript',
            ...normalizedControlProperties,
            base,
            ...(baseProperties ? { baseProperties } : {}),
            subScript,
            ...(subScriptProperties ? { subScriptProperties } : {}),
          }
        : null;
    }
    if (source.type === 'subSuperScript') {
      const alignScripts =
        source.alignScripts === undefined
          ? false
          : typeof source.alignScripts === 'boolean'
            ? source.alignScripts
            : null;
      const base = normalizeMathArgument(source.base, state);
      const subScript = normalizeMathArgument(source.subScript, state);
      const superScript = normalizeMathArgument(source.superScript, state);
      const baseProperties = normalizeMathArgumentProperties(
        source.baseProperties,
      );
      const subScriptProperties = normalizeMathArgumentProperties(
        source.subScriptProperties,
      );
      const superScriptProperties = normalizeMathArgumentProperties(
        source.superScriptProperties,
      );
      return alignScripts !== null &&
        base &&
        subScript &&
        superScript &&
        baseProperties !== null &&
        subScriptProperties !== null &&
        superScriptProperties !== null
        ? {
            type: 'subSuperScript',
            ...normalizedControlProperties,
            ...(alignScripts ? { alignScripts } : {}),
            base,
            ...(baseProperties ? { baseProperties } : {}),
            subScript,
            ...(subScriptProperties ? { subScriptProperties } : {}),
            superScript,
            ...(superScriptProperties ? { superScriptProperties } : {}),
          }
        : null;
    }
    if (source.type === 'preSubSuperScript') {
      const base = normalizeMathArgument(source.base, state);
      const subScript = normalizeMathArgument(source.subScript, state);
      const superScript = normalizeMathArgument(source.superScript, state);
      const baseProperties = normalizeMathArgumentProperties(
        source.baseProperties,
      );
      const subScriptProperties = normalizeMathArgumentProperties(
        source.subScriptProperties,
      );
      const superScriptProperties = normalizeMathArgumentProperties(
        source.superScriptProperties,
      );
      return base &&
        subScript &&
        superScript &&
        baseProperties !== null &&
        subScriptProperties !== null &&
        superScriptProperties !== null
        ? {
            type: 'preSubSuperScript',
            ...normalizedControlProperties,
            base,
            ...(baseProperties ? { baseProperties } : {}),
            subScript,
            ...(subScriptProperties ? { subScriptProperties } : {}),
            superScript,
            ...(superScriptProperties ? { superScriptProperties } : {}),
          }
        : null;
    }
    if (source.type === 'lowerLimit' || source.type === 'upperLimit') {
      const base = normalizeMathArgument(source.base, state);
      const limit = normalizeMathArgument(source.limit, state);
      const baseProperties = normalizeMathArgumentProperties(
        source.baseProperties,
      );
      const limitProperties = normalizeMathArgumentProperties(
        source.limitProperties,
      );
      return base &&
        limit &&
        baseProperties !== null &&
        limitProperties !== null
        ? {
            type: source.type,
            ...normalizedControlProperties,
            base,
            ...(baseProperties ? { baseProperties } : {}),
            limit,
            ...(limitProperties ? { limitProperties } : {}),
          }
        : null;
    }
    if (source.type === 'radical') {
      const children = normalizeMathArgument(source.children, state);
      const degree =
        source.degree === undefined
          ? undefined
          : normalizeMathArgument(source.degree, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      const degreeProperties = normalizeMathArgumentProperties(
        source.degreeProperties,
      );
      return children &&
        degree !== null &&
        childrenProperties !== null &&
        degreeProperties !== null
        ? {
            type: 'radical',
            ...normalizedControlProperties,
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
            ...(degree?.length ? { degree } : {}),
            ...(degreeProperties ? { degreeProperties } : {}),
          }
        : null;
    }
    if (source.type === 'function') {
      const name = normalizeMathArgument(source.name, state);
      const children = normalizeMathArgument(source.children, state);
      const nameProperties = normalizeMathArgumentProperties(
        source.nameProperties,
      );
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return name &&
        children &&
        nameProperties !== null &&
        childrenProperties !== null
        ? {
            type: 'function',
            ...normalizedControlProperties,
            name,
            ...(nameProperties ? { nameProperties } : {}),
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'nary') {
      const operator = NARY_OPERATORS.has(
        source.operator as WorkDocumentEquationNaryOperator,
      )
        ? (source.operator as WorkDocumentEquationNaryOperator)
        : null;
      const limitLocation = LIMIT_LOCATIONS.has(
        source.limitLocation as WorkDocumentEquationLimitLocation,
      )
        ? (source.limitLocation as WorkDocumentEquationLimitLocation)
        : null;
      const grow = source.grow;
      const children = normalizeMathArgument(source.children, state);
      const subScript =
        source.subScript === undefined
          ? undefined
          : normalizeMathArgument(source.subScript, state);
      const superScript =
        source.superScript === undefined
          ? undefined
          : normalizeMathArgument(source.superScript, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      const subScriptProperties = normalizeMathArgumentProperties(
        source.subScriptProperties,
      );
      const superScriptProperties = normalizeMathArgumentProperties(
        source.superScriptProperties,
      );
      return operator &&
        limitLocation &&
        (grow === undefined || typeof grow === 'boolean') &&
        children &&
        subScript !== null &&
        superScript !== null &&
        childrenProperties !== null &&
        subScriptProperties !== null &&
        superScriptProperties !== null
        ? {
            type: 'nary',
            ...normalizedControlProperties,
            operator,
            limitLocation,
            ...(grow === true ? { grow: true } : {}),
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
            ...(subScript?.length ? { subScript } : {}),
            ...(subScriptProperties ? { subScriptProperties } : {}),
            ...(superScript?.length ? { superScript } : {}),
            ...(superScriptProperties ? { superScriptProperties } : {}),
          }
        : null;
    }
    if (source.type === 'accent') {
      const character = accentCharacter(source.character);
      const children = normalizeMathArgument(source.children, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return character && children && childrenProperties !== null
        ? {
            type: 'accent',
            ...normalizedControlProperties,
            character,
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'bar') {
      const position = BAR_POSITIONS.has(
        source.position as WorkDocumentEquationBarPosition,
      )
        ? (source.position as WorkDocumentEquationBarPosition)
        : null;
      const children = normalizeMathArgument(source.children, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return position && children && childrenProperties !== null
        ? {
            type: 'bar',
            ...normalizedControlProperties,
            position,
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'groupCharacter') {
      const character = mathCharacter(source.character);
      const position = BAR_POSITIONS.has(
        source.position as WorkDocumentEquationBarPosition,
      )
        ? (source.position as WorkDocumentEquationBarPosition)
        : null;
      const verticalJustification = BAR_POSITIONS.has(
        source.verticalJustification as WorkDocumentEquationBarPosition,
      )
        ? (source.verticalJustification as WorkDocumentEquationBarPosition)
        : null;
      const children = normalizeMathArgument(source.children, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return character !== null &&
        position &&
        verticalJustification &&
        children &&
        childrenProperties !== null
        ? {
            type: 'groupCharacter',
            ...normalizedControlProperties,
            character,
            position,
            verticalJustification,
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'phantom') {
      if (
        typeof source.show !== 'boolean' ||
        typeof source.zeroWidth !== 'boolean' ||
        typeof source.zeroAscent !== 'boolean' ||
        typeof source.zeroDescent !== 'boolean' ||
        typeof source.transparent !== 'boolean'
      ) {
        return null;
      }
      const children = normalizeMathArgument(source.children, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return children && childrenProperties !== null
        ? {
            type: 'phantom',
            ...normalizedControlProperties,
            show: source.show,
            zeroWidth: source.zeroWidth,
            zeroAscent: source.zeroAscent,
            zeroDescent: source.zeroDescent,
            transparent: source.transparent,
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'borderBox') {
      if (
        typeof source.hideTop !== 'boolean' ||
        typeof source.hideBottom !== 'boolean' ||
        typeof source.hideLeft !== 'boolean' ||
        typeof source.hideRight !== 'boolean' ||
        typeof source.strikeHorizontal !== 'boolean' ||
        typeof source.strikeVertical !== 'boolean' ||
        typeof source.strikeBottomLeftToTopRight !== 'boolean' ||
        typeof source.strikeTopLeftToBottomRight !== 'boolean'
      ) {
        return null;
      }
      const children = normalizeMathArgument(source.children, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return children && childrenProperties !== null
        ? {
            type: 'borderBox',
            ...normalizedControlProperties,
            hideTop: source.hideTop,
            hideBottom: source.hideBottom,
            hideLeft: source.hideLeft,
            hideRight: source.hideRight,
            strikeHorizontal: source.strikeHorizontal,
            strikeVertical: source.strikeVertical,
            strikeBottomLeftToTopRight: source.strikeBottomLeftToTopRight,
            strikeTopLeftToBottomRight: source.strikeTopLeftToBottomRight,
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'box') {
      const manualBreak =
        source.manualBreak === undefined
          ? undefined
          : normalizeManualBreak(source.manualBreak);
      const children = normalizeMathArgument(source.children, state);
      const childrenProperties = normalizeMathArgumentProperties(
        source.childrenProperties,
      );
      return typeof source.operatorEmulator === 'boolean' &&
        typeof source.noBreak === 'boolean' &&
        typeof source.differential === 'boolean' &&
        typeof source.alignment === 'boolean' &&
        manualBreak !== null &&
        children &&
        childrenProperties !== null
        ? {
            type: 'box',
            ...normalizedControlProperties,
            operatorEmulator: source.operatorEmulator,
            noBreak: source.noBreak,
            differential: source.differential,
            alignment: source.alignment,
            ...(manualBreak ? { manualBreak } : {}),
            children,
            ...(childrenProperties ? { childrenProperties } : {}),
          }
        : null;
    }
    if (source.type === 'matrix') {
      const baseAlignment = MATRIX_BASE_ALIGNMENTS.has(
        source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment,
      )
        ? (source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment)
        : null;
      const spacing =
        source.spacing === undefined
          ? undefined
          : normalizeMatrixSpacing(source.spacing);
      if (
        !baseAlignment ||
        typeof source.placeholdersHidden !== 'boolean' ||
        spacing === null ||
        !Array.isArray(source.rows) ||
        source.rows.length === 0 ||
        source.rows.length > MAX_MATRIX_ROWS ||
        !Array.isArray(source.columnAlignments)
      ) {
        return null;
      }
      const columnCount = Array.isArray(source.rows[0])
        ? source.rows[0].length
        : 0;
      if (
        columnCount === 0 ||
        columnCount > MAX_MATRIX_COLUMNS ||
        source.rows.length * columnCount > MAX_MATRIX_CELLS ||
        source.columnAlignments.length !== columnCount
      ) {
        return null;
      }
      const columnAlignments = source.columnAlignments.map((alignment) =>
        MATRIX_ALIGNMENTS.has(alignment as WorkDocumentEquationMatrixAlignment)
          ? (alignment as WorkDocumentEquationMatrixAlignment)
          : null,
      );
      if (
        !columnAlignments.every(
          (alignment): alignment is WorkDocumentEquationMatrixAlignment =>
            alignment !== null,
        )
      ) {
        return null;
      }
      const rows: WorkDocumentEquationExpression[][][] = [];
      for (const row of source.rows) {
        if (!Array.isArray(row) || row.length !== columnCount) return null;
        const normalizedRow: WorkDocumentEquationExpression[][] = [];
        for (const cell of row) {
          const normalizedCell = normalizeMathArgument(cell, state);
          if (!normalizedCell) return null;
          normalizedRow.push(normalizedCell);
        }
        rows.push(normalizedRow);
      }
      const cellProperties = normalizeMatrixArgumentPropertySlots(
        source.cellProperties,
        source.rows,
      );
      if (cellProperties === null) return null;
      return {
        type: 'matrix',
        ...normalizedControlProperties,
        baseAlignment,
        placeholdersHidden: source.placeholdersHidden,
        columnAlignments,
        ...(spacing ? { spacing } : {}),
        rows,
        ...(cellProperties ? { cellProperties } : {}),
      };
    }
    if (source.type === 'equationArray') {
      const baseAlignment = MATRIX_BASE_ALIGNMENTS.has(
        source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment,
      )
        ? (source.baseAlignment as WorkDocumentEquationMatrixBaseAlignment)
        : null;
      const rowSpacingRule = EQUATION_SPACING_RULES.has(
        source.rowSpacingRule as WorkDocumentEquationRowSpacingRule,
      )
        ? (source.rowSpacingRule as WorkDocumentEquationRowSpacingRule)
        : null;
      if (
        !baseAlignment ||
        typeof source.maximumDistribution !== 'boolean' ||
        typeof source.objectDistribution !== 'boolean' ||
        !rowSpacingRule ||
        typeof source.rowSpacing !== 'number' ||
        !Number.isInteger(source.rowSpacing) ||
        source.rowSpacing < 0 ||
        source.rowSpacing > 65_535 ||
        !Array.isArray(source.rows) ||
        source.rows.length === 0 ||
        source.rows.length > MAX_EQUATION_ARRAY_ROWS
      ) {
        return null;
      }
      const rows: WorkDocumentEquationExpression[][] = [];
      state.equationArrayDepth += 1;
      try {
        for (const row of source.rows) {
          const normalizedRow = normalizeMathArgument(row, state);
          if (!normalizedRow) return null;
          rows.push(normalizedRow);
        }
      } finally {
        state.equationArrayDepth -= 1;
      }
      const rowProperties = normalizeMathArgumentPropertySlots(
        source.rowProperties,
        rows.length,
      );
      if (rowProperties === null) return null;
      return {
        type: 'equationArray',
        ...normalizedControlProperties,
        baseAlignment,
        maximumDistribution: source.maximumDistribution,
        objectDistribution: source.objectDistribution,
        rowSpacingRule,
        rowSpacing: source.rowSpacing,
        rows,
        ...(rowProperties ? { rowProperties } : {}),
      };
    }
    if (source.type === 'delimiter') {
      const opening = mathCharacter(source.opening);
      const closing = mathCharacter(source.closing);
      const separator = mathCharacter(source.separator);
      const grow = source.grow;
      const shape =
        source.shape === undefined
          ? undefined
          : DELIMITER_SHAPES.has(
                source.shape as WorkDocumentEquationDelimiterShape,
              )
            ? (source.shape as WorkDocumentEquationDelimiterShape)
            : null;
      if (
        opening === null ||
        closing === null ||
        separator === null ||
        (grow !== undefined && typeof grow !== 'boolean') ||
        shape === null ||
        !Array.isArray(source.arguments) ||
        source.arguments.length === 0 ||
        source.arguments.length > MAX_DELIMITER_ARGUMENTS
      ) {
        return null;
      }
      const arguments_ = source.arguments.map((argument) =>
        normalizeMathArgument(argument, state),
      );
      const argumentProperties = normalizeMathArgumentPropertySlots(
        source.argumentProperties,
        source.arguments.length,
      );
      return arguments_.every(
        (argument): argument is WorkDocumentEquationExpression[] =>
          Boolean(argument),
      ) && argumentProperties !== null
        ? {
            type: 'delimiter',
            ...normalizedControlProperties,
            opening,
            closing,
            separator,
            ...(grow === false ? { grow: false } : {}),
            ...(shape === 'match' ? { shape } : {}),
            arguments: arguments_,
            ...(argumentProperties ? { argumentProperties } : {}),
          }
        : null;
    }
    return null;
  } finally {
    state.depth -= 1;
  }
}

function expressionListText(
  expressions: readonly WorkDocumentEquationExpression[],
  hideAlignmentMarkers = false,
): string {
  return expressions
    .map((expression) => expressionText(expression, hideAlignmentMarkers))
    .join('');
}

function expressionText(
  expression: WorkDocumentEquationExpression,
  hideAlignmentMarkers = false,
): string {
  if (expression.type === 'run') {
    const text = hideAlignmentMarkers
      ? expression.text.replaceAll('&', '')
      : expression.text;
    const properties = [
      expression.literal ? 'literal' : '',
      expression.normalText ? 'normal-text' : '',
      expression.script ? `script=${expression.script}` : '',
      expression.style ? `style=${expression.style}` : '',
      expression.manualBreak
        ? expression.manualBreak.alignmentAt
          ? `break@${expression.manualBreak.alignmentAt}`
          : 'break'
        : '',
      expression.alignment ? 'alignment' : '',
    ].filter(Boolean);
    return properties.length ? `run(${properties.join(',')};${text})` : text;
  }
  if (expression.type === 'fraction') {
    return `(${expressionListText(
      expression.numerator,
      hideAlignmentMarkers,
    )})/(${expressionListText(expression.denominator, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'superscript') {
    return `${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )}^(${expressionListText(expression.superScript, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'subscript') {
    return `${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )}_(${expressionListText(expression.subScript, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'subSuperScript') {
    return `${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )}_(${expressionListText(
      expression.subScript,
      hideAlignmentMarkers,
    )})^(${expressionListText(expression.superScript, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'radical') {
    const body = expressionListText(expression.children, hideAlignmentMarkers);
    return expression.degree
      ? `root(${expressionListText(
          expression.degree,
          hideAlignmentMarkers,
        )};${body})`
      : `sqrt(${body})`;
  }
  if (expression.type === 'function') {
    return `${expressionListText(
      expression.name,
      hideAlignmentMarkers,
    )}(${expressionListText(expression.children, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'nary') {
    const text = `${expression.operator}${
      expression.subScript
        ? `_(${expressionListText(expression.subScript, hideAlignmentMarkers)})`
        : ''
    }${
      expression.superScript
        ? `^(${expressionListText(
            expression.superScript,
            hideAlignmentMarkers,
          )})`
        : ''
    } ${expressionListText(expression.children, hideAlignmentMarkers)}`;
    return expression.grow ? `nary(grow;${text})` : text;
  }
  if (expression.type === 'accent') {
    const codePoint = expression.character.codePointAt(0);
    const label = codePoint?.toString(16).toUpperCase().padStart(4, '0');
    return `accent(U+${label};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'preSubSuperScript') {
    const subScript = expressionListText(
      expression.subScript,
      hideAlignmentMarkers,
    );
    const superScript = expressionListText(
      expression.superScript,
      hideAlignmentMarkers,
    );
    return `pre-scripts(sub=${subScript || 'none'};sup=${
      superScript || 'none'
    };base=${expressionListText(expression.base, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'lowerLimit' || expression.type === 'upperLimit') {
    return `${expression.type === 'lowerLimit' ? 'lower-limit' : 'upper-limit'}(${expressionListText(
      expression.base,
      hideAlignmentMarkers,
    )};${expressionListText(expression.limit, hideAlignmentMarkers)})`;
  }
  if (expression.type === 'bar') {
    const body = expressionListText(expression.children, hideAlignmentMarkers);
    return expression.position === 'top'
      ? `overbar(${body})`
      : `underbar(${body})`;
  }
  if (expression.type === 'groupCharacter') {
    const codePoint = expression.character.codePointAt(0);
    const character =
      codePoint === undefined
        ? 'none'
        : `U+${codePoint.toString(16).toUpperCase().padStart(4, '0')}`;
    return `group-character(${character};position=${expression.position};baseline=${expression.verticalJustification};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'phantom') {
    const properties = [
      expression.show ? 'visible' : 'hidden',
      expression.zeroWidth ? 'zero-width' : '',
      expression.zeroAscent ? 'zero-ascent' : '',
      expression.zeroDescent ? 'zero-descent' : '',
      expression.transparent ? 'transparent' : '',
    ].filter(Boolean);
    return `phantom(${properties.join(',')};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'borderBox') {
    return `borderbox(${borderBoxNotation(expression)};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'box') {
    const properties = [
      expression.operatorEmulator ? 'operator' : '',
      expression.noBreak ? 'no-break' : '',
      expression.differential ? 'differential' : '',
      expression.manualBreak
        ? expression.manualBreak.alignmentAt
          ? `break@${expression.manualBreak.alignmentAt}`
          : 'break'
        : '',
      expression.alignment ? 'alignment' : '',
    ].filter(Boolean);
    return `box(${properties.join(',') || 'default'};${expressionListText(
      expression.children,
      hideAlignmentMarkers,
    )})`;
  }
  if (expression.type === 'matrix') {
    const spacing = expression.spacing
      ? `row-spacing=${expression.spacing.rowSpacingRule}:${expression.spacing.rowSpacing},column-gap=${expression.spacing.columnGapRule}:${expression.spacing.columnGap},minimum-column-width=${expression.spacing.minimumColumnWidthTwips}twip;`
      : '';
    return `matrix(${spacing}${expression.rows
      .map((row) =>
        row
          .map((cell) => expressionListText(cell, hideAlignmentMarkers))
          .join(','),
      )
      .join(';')})`;
  }
  if (expression.type === 'equationArray') {
    const properties = [
      expression.baseAlignment,
      expression.maximumDistribution ? 'max-distribution' : '',
      expression.objectDistribution ? 'object-distribution' : '',
      `spacing=${expression.rowSpacingRule}:${expression.rowSpacing}`,
    ].filter(Boolean);
    return `equation-array(${properties.join(',')};${expression.rows
      .map((row) => expressionListText(row, true))
      .join(';')})`;
  }
  const text = `${expression.opening}${expression.arguments
    .map((argument) => expressionListText(argument, hideAlignmentMarkers))
    .join(expression.separator)}${expression.closing}`;
  const properties = [
    expression.grow === false ? 'grow=false' : '',
    expression.shape === 'match' ? 'shape=match' : '',
  ].filter(Boolean);
  return properties.length
    ? `delimiter(${properties.join(',')};${text})`
    : text;
}

function equationMathMl(equation: WorkDocumentEquation): DOMOutputSpec {
  return domSpec(
    'math',
    {
      display: equation.display === 'block' ? 'block' : 'inline',
      xmlns: MATHML_NAMESPACE,
    },
    [mathRow(equation.children)],
  );
}

function createEquationDom(document: Document, spec: DOMOutputSpec): Element {
  if (!Array.isArray(spec) || typeof spec[0] !== 'string') {
    throw new Error('The document equation rendering model is invalid.');
  }
  const element = document.createElementNS(MATHML_NAMESPACE, spec[0]);
  const attributes = spec[1];
  if (
    attributes &&
    typeof attributes === 'object' &&
    !Array.isArray(attributes)
  ) {
    for (const [name, value] of Object.entries(attributes)) {
      if (typeof value === 'string') element.setAttribute(name, value);
    }
  }
  for (const child of spec.slice(2)) {
    if (typeof child === 'string') {
      element.append(document.createTextNode(child));
    } else if (Array.isArray(child) && typeof child[0] === 'string') {
      element.append(
        createEquationDom(document, child as unknown as DOMOutputSpec),
      );
    }
  }
  return element;
}

function mathRow(
  expressions: readonly WorkDocumentEquationExpression[],
  alignmentState?: EquationArrayAlignmentState,
): DOMOutputSpec {
  return domSpec(
    'mrow',
    {},
    expressions.map((expression) =>
      expressionMathMl(expression, alignmentState),
    ),
  );
}

function wordSizedMathArgument(
  expressions: readonly WorkDocumentEquationExpression[],
  properties: WorkDocumentEquationArgumentProperties | undefined,
  alignmentState?: EquationArrayAlignmentState,
): DOMOutputSpec {
  const row = mathRow(expressions, alignmentState);
  if (!expressions.length || properties?.size === undefined) return row;
  const relativeScriptLevel =
    properties.size < 0
      ? `+${Math.abs(properties.size)}`
      : `-${properties.size}`;
  return domSpec('mstyle', { scriptlevel: relativeScriptLevel }, [row]);
}

function expressionMathMl(
  expression: WorkDocumentEquationExpression,
  alignmentState?: EquationArrayAlignmentState,
): DOMOutputSpec {
  if (expression.type === 'run') {
    const mathVariant = runMathVariant(expression);
    const attributes: Record<string, string> = {
      ...(mathVariant ? { mathvariant: mathVariant } : {}),
      ...wordRunMathMlAttributes(expression),
    };
    if (!alignmentState) {
      return domSpec('mtext', attributes, [expression.text]);
    }
    const children: DOMOutputSpec[] = [];
    if (!alignmentState.started) {
      children.push(domSpec('maligngroup', {}, []));
      alignmentState.started = true;
    }
    const parts = expression.text.split('&');
    parts.forEach((part, index) => {
      if (part) children.push(domSpec('mtext', attributes, [part]));
      if (index === parts.length - 1) return;
      alignmentState.markerIndex += 1;
      children.push(
        domSpec(
          alignmentState.markerIndex % 2 === 1 ? 'malignmark' : 'maligngroup',
          {},
          [],
        ),
      );
    });
    return domSpec('mrow', {}, children);
  }
  if (expression.type === 'fraction') {
    if (expression.fractionType === 'linear') {
      return domSpec('mrow', {}, [
        mathRow(expression.numerator, alignmentState),
        domSpec('mo', controlMathMlAttributes(expression, '/'), ['/']),
        mathRow(expression.denominator, alignmentState),
      ]);
    }
    return domSpec(
      'mfrac',
      {
        ...(expression.fractionType === 'noBar' ? { linethickness: '0' } : {}),
        ...(expression.fractionType === 'skewed' ? { bevelled: 'true' } : {}),
      },
      [
        mathRow(expression.numerator, alignmentState),
        mathRow(expression.denominator, alignmentState),
      ],
    );
  }
  if (expression.type === 'superscript') {
    return domSpec('msup', {}, [
      mathRow(expression.base, alignmentState),
      wordSizedMathArgument(
        expression.superScript,
        expression.superScriptProperties,
        alignmentState,
      ),
    ]);
  }
  if (expression.type === 'subscript') {
    return domSpec('msub', {}, [
      mathRow(expression.base, alignmentState),
      wordSizedMathArgument(
        expression.subScript,
        expression.subScriptProperties,
        alignmentState,
      ),
    ]);
  }
  if (expression.type === 'subSuperScript') {
    return domSpec('msubsup', {}, [
      mathRow(expression.base, alignmentState),
      wordSizedMathArgument(
        expression.subScript,
        expression.subScriptProperties,
        alignmentState,
      ),
      wordSizedMathArgument(
        expression.superScript,
        expression.superScriptProperties,
        alignmentState,
      ),
    ]);
  }
  if (expression.type === 'preSubSuperScript') {
    return domSpec('mmultiscripts', {}, [
      mathRow(expression.base, alignmentState),
      domSpec('mprescripts', {}, []),
      expression.subScript.length
        ? wordSizedMathArgument(
            expression.subScript,
            expression.subScriptProperties,
            alignmentState,
          )
        : domSpec('none', {}, []),
      expression.superScript.length
        ? wordSizedMathArgument(
            expression.superScript,
            expression.superScriptProperties,
            alignmentState,
          )
        : domSpec('none', {}, []),
    ]);
  }
  if (expression.type === 'lowerLimit' || expression.type === 'upperLimit') {
    return domSpec(
      expression.type === 'lowerLimit' ? 'munder' : 'mover',
      expression.type === 'lowerLimit'
        ? { accentunder: 'false' }
        : { accent: 'false' },
      [
        mathRow(expression.base, alignmentState),
        wordSizedMathArgument(
          expression.limit,
          expression.limitProperties,
          alignmentState,
        ),
      ],
    );
  }
  if (expression.type === 'radical') {
    return expression.degree
      ? domSpec('mroot', {}, [
          mathRow(expression.children, alignmentState),
          wordSizedMathArgument(
            expression.degree,
            expression.degreeProperties,
            alignmentState,
          ),
        ])
      : domSpec('msqrt', {}, [mathRow(expression.children, alignmentState)]);
  }
  if (expression.type === 'function') {
    return domSpec('mrow', {}, [
      mathRow(expression.name, alignmentState),
      domSpec('mo', controlMathMlAttributes(expression, '\u2061'), ['\u2061']),
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'nary') {
    const operator = domSpec(
      'mo',
      {
        ...controlMathMlAttributes(expression, expression.operator),
        stretchy: expression.grow ? 'true' : 'false',
      },
      [expression.operator],
    );
    let decorated = operator;
    if (expression.subScript && expression.superScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'munderover' : 'msubsup',
        {},
        [
          operator,
          wordSizedMathArgument(
            expression.subScript,
            expression.subScriptProperties,
            alignmentState,
          ),
          wordSizedMathArgument(
            expression.superScript,
            expression.superScriptProperties,
            alignmentState,
          ),
        ],
      );
    } else if (expression.subScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'munder' : 'msub',
        {},
        [
          operator,
          wordSizedMathArgument(
            expression.subScript,
            expression.subScriptProperties,
            alignmentState,
          ),
        ],
      );
    } else if (expression.superScript) {
      decorated = domSpec(
        expression.limitLocation === 'underOver' ? 'mover' : 'msup',
        {},
        [
          operator,
          wordSizedMathArgument(
            expression.superScript,
            expression.superScriptProperties,
            alignmentState,
          ),
        ],
      );
    }
    return domSpec('mrow', {}, [
      decorated,
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'accent') {
    const character =
      MATHML_ACCENT_CHARACTERS.get(expression.character) ??
      expression.character;
    return domSpec('mover', { accent: 'true' }, [
      mathRow(expression.children, alignmentState),
      domSpec('mo', controlMathMlAttributes(expression, character), [
        character,
      ]),
    ]);
  }
  if (expression.type === 'bar') {
    return domSpec(
      expression.position === 'top' ? 'mover' : 'munder',
      expression.position === 'top'
        ? { accent: 'false' }
        : { accentunder: 'false' },
      [
        mathRow(expression.children, alignmentState),
        domSpec('mo', controlMathMlAttributes(expression, '\u00af'), [
          '\u00af',
        ]),
      ],
    );
  }
  if (expression.type === 'groupCharacter') {
    return domSpec(
      expression.position === 'top' ? 'mover' : 'munder',
      expression.position === 'top'
        ? { accent: 'false' }
        : { accentunder: 'false' },
      [
        wordSizedMathArgument(
          expression.children,
          expression.childrenProperties,
          alignmentState,
        ),
        domSpec(
          'mo',
          controlMathMlAttributes(expression, expression.character),
          [expression.character],
        ),
      ],
    );
  }
  if (expression.type === 'phantom') {
    const attributes = {
      ...(expression.zeroWidth ? { width: '0in' } : {}),
      ...(expression.zeroAscent ? { height: '0in' } : {}),
      ...(expression.zeroDescent ? { depth: '0in' } : {}),
    };
    const body = mathRow(expression.children, alignmentState);
    const padded = domSpec('mpadded', attributes, [body]);
    if (expression.show) return padded;
    return domSpec('mphantom', {}, [
      Object.keys(attributes).length ? padded : body,
    ]);
  }
  if (expression.type === 'borderBox') {
    return domSpec('menclose', { notation: borderBoxNotation(expression) }, [
      mathRow(expression.children, alignmentState),
    ]);
  }
  if (expression.type === 'box') {
    return domSpec('mpadded', {}, [
      wordSizedMathArgument(
        expression.children,
        expression.childrenProperties,
        alignmentState,
      ),
    ]);
  }
  if (expression.type === 'matrix') {
    return domSpec(
      'mtable',
      {
        align: expression.baseAlignment,
        columnalign: expression.columnAlignments.join(' '),
        ...(expression.spacing
          ? {
              rowspacing: equationRowSpacing(
                expression.spacing.rowSpacingRule,
                expression.spacing.rowSpacing,
              ),
              columnspacing: matrixColumnGap(
                expression.spacing.columnGapRule,
                expression.spacing.columnGap,
              ),
            }
          : {}),
      },
      expression.rows.map((row) =>
        domSpec(
          'mtr',
          {},
          row.map((cell) =>
            domSpec('mtd', {}, [mathRow(cell, alignmentState)]),
          ),
        ),
      ),
    );
  }
  if (expression.type === 'equationArray') {
    return domSpec(
      'mtable',
      {
        align: expression.baseAlignment,
        rowspacing: equationRowSpacing(
          expression.rowSpacingRule,
          expression.rowSpacing,
        ),
      },
      expression.rows.map((row) => {
        const rowAlignmentState: EquationArrayAlignmentState = {
          markerIndex: 0,
          started: false,
        };
        return domSpec('mtr', {}, [
          domSpec('mtd', {}, [mathRow(row, rowAlignmentState)]),
        ]);
      }),
    );
  }
  const children: Array<DOMOutputSpec | string> = [];
  const sizingAttributes = {
    stretchy: expression.grow === false ? 'false' : 'true',
    ...(expression.grow === false
      ? {}
      : { symmetric: expression.shape === 'match' ? 'false' : 'true' }),
  };
  if (expression.opening)
    children.push(
      domSpec(
        'mo',
        {
          fence: 'true',
          ...sizingAttributes,
          ...controlMathMlAttributes(expression, expression.opening),
        },
        [expression.opening],
      ),
    );
  expression.arguments.forEach((argument, index) => {
    if (index > 0 && expression.separator) {
      children.push(
        domSpec(
          'mo',
          {
            separator: 'true',
            ...sizingAttributes,
            ...controlMathMlAttributes(expression, expression.separator),
          },
          [expression.separator],
        ),
      );
    }
    children.push(mathRow(argument, alignmentState));
  });
  if (expression.closing)
    children.push(
      domSpec(
        'mo',
        {
          fence: 'true',
          ...sizingAttributes,
          ...controlMathMlAttributes(expression, expression.closing),
        },
        [expression.closing],
      ),
    );
  return domSpec('mrow', {}, children);
}

function equationRowSpacing(
  rule: WorkDocumentEquationSpacingRule,
  value: number,
): string {
  if (rule === 'oneAndHalf') return '1.5em';
  if (rule === 'double') return '2em';
  if (rule === 'exact') {
    return `${value}pt`;
  }
  if (rule === 'multiple') {
    return `${value / 2}em`;
  }
  return '1em';
}

function matrixColumnGap(
  rule: WorkDocumentEquationSpacingRule,
  value: number,
): string {
  if (rule === 'exact') return `${value / 20}pt`;
  return equationRowSpacing(rule, value);
}

function domSpec(
  name: string,
  attributes: Record<string, string>,
  children: readonly (DOMOutputSpec | string)[],
): DOMOutputSpec {
  return [name, attributes, ...children] as DOMOutputSpec;
}

function mathCharacter(source: unknown): string | null {
  if (typeof source !== 'string') return null;
  if (source === '') return '';
  const characters = Array.from(source);
  if (characters.length !== 1 || /[\p{Cc}\p{Cs}]/u.test(source)) return null;
  return source;
}

function accentCharacter(source: unknown): string | null {
  if (typeof source !== 'string' || Array.from(source).length !== 1) {
    return null;
  }
  const codePoint = source.codePointAt(0);
  return codePoint !== undefined &&
    ((codePoint >= 0x0300 && codePoint <= 0x036f) ||
      (codePoint >= 0x20d0 && codePoint <= 0x20ef))
    ? source
    : null;
}

function normalizeEquationControlRevision(
  source: unknown,
  allowedKinds: ReadonlySet<
    WorkDocumentEquationControlRevision['kind']
  > = CONTROL_REVISION_KINDS,
): WorkDocumentEquationControlRevision | null {
  if (
    !isRecordWithKeys(source, CONTROL_REVISION_KEYS) ||
    !allowedKinds.has(
      source.kind as WorkDocumentEquationControlRevision['kind'],
    ) ||
    typeof source.id !== 'number' ||
    !Number.isInteger(source.id) ||
    source.id < 0 ||
    source.id > MAX_EQUATION_CONTROL_REVISION_ID
  ) {
    return null;
  }
  const author = normalizedEquationWordString(
    source.author,
    MAX_EQUATION_CONTROL_REVISION_AUTHOR_LENGTH,
  );
  const date =
    source.date === undefined
      ? undefined
      : normalizedEquationControlRevisionDate(source.date);
  const dateUtc =
    source.dateUtc === undefined
      ? undefined
      : normalizedEquationControlRevisionUtcDate(source.dateUtc);
  if (!author || date === null || dateUtc === null) return null;

  const kind = source.kind as WorkDocumentEquationControlRevision['kind'];
  if (kind === 'deletion') {
    return source.child === undefined
      ? {
          kind,
          id: source.id,
          author,
          ...(date ? { date } : {}),
          ...(dateUtc ? { dateUtc } : {}),
        }
      : null;
  }
  const child =
    source.child === undefined
      ? undefined
      : normalizeEquationControlRevision(
          source.child,
          kind === 'insertion'
            ? CONTROL_REVISION_INSERTION_CHILD_KINDS
            : CONTROL_REVISION_MOVE_CHILD_KINDS,
        );
  if (child === null) return null;
  return {
    kind,
    id: source.id,
    author,
    ...(date ? { date } : {}),
    ...(dateUtc ? { dateUtc } : {}),
    ...(child ? { child } : {}),
  } as WorkDocumentEquationControlRevision;
}

function normalizedEquationControlRevisionDate(source: unknown): string | null {
  const value = normalizedEquationWordString(
    source,
    MAX_EQUATION_CONTROL_REVISION_DATE_LENGTH,
  );
  const match = value?.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(?:Z|([+-])(\d{2}):(\d{2}))?$/u,
  );
  if (!value || !match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const offsetHour = match[8] === undefined ? 0 : Number(match[8]);
  const offsetMinute = match[9] === undefined ? 0 : Number(match[9]);
  const leapYear = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth = [
    31,
    leapYear ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];
  return year >= 1 &&
    month >= 1 &&
    month <= 12 &&
    day >= 1 &&
    day <= (daysInMonth[month - 1] ?? 0) &&
    hour <= 23 &&
    minute <= 59 &&
    second <= 59 &&
    offsetHour <= 14 &&
    offsetMinute <= 59 &&
    (offsetHour !== 14 || offsetMinute === 0)
    ? value
    : null;
}

function normalizedEquationControlRevisionUtcDate(
  source: unknown,
): string | null {
  const date = normalizedEquationControlRevisionDate(source);
  return date?.endsWith('Z') ? date : null;
}

function normalizeEquationWordRunProperties(
  source: unknown,
): WorkDocumentEquationWordRunProperties | null | undefined {
  if (!isRecordWithKeys(source, WORD_RUN_PROPERTY_KEYS)) return null;
  const fonts =
    source.fonts === undefined
      ? undefined
      : normalizeEquationWordRunFonts(source.fonts);
  const color =
    source.color === undefined
      ? undefined
      : normalizeEquationWordColor(source.color);
  const underline =
    source.underline === undefined
      ? undefined
      : normalizeEquationWordUnderline(source.underline);
  const border =
    source.border === undefined
      ? undefined
      : normalizeEquationWordRunBorder(source.border);
  const shading =
    source.shading === undefined
      ? undefined
      : normalizeEquationWordShading(source.shading);
  const fitText =
    source.fitText === undefined
      ? undefined
      : normalizeEquationWordFitText(source.fitText);
  const verticalAlignment =
    source.verticalAlignment === undefined
      ? undefined
      : WORD_VERTICAL_ALIGNMENTS.has(
            source.verticalAlignment as WorkDocumentEquationWordVerticalAlignment,
          )
        ? (source.verticalAlignment as WorkDocumentEquationWordVerticalAlignment)
        : null;
  const emphasisMark =
    source.emphasisMark === undefined
      ? undefined
      : WORD_EMPHASIS_MARKS.has(
            source.emphasisMark as WorkDocumentEquationWordEmphasisMark,
          )
        ? (source.emphasisMark as WorkDocumentEquationWordEmphasisMark)
        : null;
  const languages =
    source.languages === undefined
      ? undefined
      : normalizeEquationWordLanguages(source.languages);
  const eastAsianLayout =
    source.eastAsianLayout === undefined
      ? undefined
      : normalizeEquationWordEastAsianLayout(source.eastAsianLayout);
  const glow =
    source.glow === undefined
      ? undefined
      : normalizeEquationWordGlow(source.glow);
  const shadowEffect =
    source.shadowEffect === undefined
      ? undefined
      : normalizeEquationWordShadowEffect(source.shadowEffect);
  const reflectionEffect =
    source.reflectionEffect === undefined
      ? undefined
      : normalizeEquationWordReflectionEffect(source.reflectionEffect);
  const textOutlineEffect =
    source.textOutlineEffect === undefined
      ? undefined
      : normalizeEquationWordTextOutlineEffect(source.textOutlineEffect);
  const textFillEffect =
    source.textFillEffect === undefined
      ? undefined
      : normalizeEquationWordTextFillEffect(source.textFillEffect);
  const scene3D =
    source.scene3D === undefined
      ? undefined
      : normalizeEquationWordScene3D(source.scene3D);
  const properties3D =
    source.properties3D === undefined
      ? undefined
      : normalizeEquationWordProperties3D(source.properties3D);
  const ligatures =
    source.ligatures === undefined
      ? undefined
      : WORD_LIGATURE_FLAGS.has(
            source.ligatures as WorkDocumentEquationWordLigatures,
          )
        ? (source.ligatures as WorkDocumentEquationWordLigatures)
        : null;
  const numberForm =
    source.numberForm === undefined
      ? undefined
      : WORD_NUMBER_FORMS.has(
            source.numberForm as WorkDocumentEquationWordNumberForm,
          )
        ? (source.numberForm as WorkDocumentEquationWordNumberForm)
        : null;
  const numberSpacing =
    source.numberSpacing === undefined
      ? undefined
      : WORD_NUMBER_SPACINGS.has(
            source.numberSpacing as WorkDocumentEquationWordNumberSpacing,
          )
        ? (source.numberSpacing as WorkDocumentEquationWordNumberSpacing)
        : null;
  const stylisticSets =
    source.stylisticSets === undefined
      ? undefined
      : normalizeEquationWordStylisticSets(source.stylisticSets);
  const characterSpacingTwips =
    source.characterSpacingTwips === undefined
      ? undefined
      : normalizeEquationInteger(
          source.characterSpacingTwips,
          -MAX_EQUATION_CHARACTER_SPACING_TWIPS,
          MAX_EQUATION_CHARACTER_SPACING_TWIPS,
        );
  const characterScalePercent =
    source.characterScalePercent === undefined
      ? undefined
      : normalizeEquationInteger(
          source.characterScalePercent,
          1,
          MAX_EQUATION_CHARACTER_SCALE_PERCENT,
        );
  const kerningThresholdHalfPoints =
    source.kerningThresholdHalfPoints === undefined
      ? undefined
      : normalizeEquationInteger(
          source.kerningThresholdHalfPoints,
          0,
          MAX_EQUATION_KERNING_THRESHOLD_HALF_POINTS,
        );
  const positionHalfPoints =
    source.positionHalfPoints === undefined
      ? undefined
      : normalizeEquationInteger(
          source.positionHalfPoints,
          MIN_EQUATION_POSITION_HALF_POINTS,
          MAX_EQUATION_POSITION_HALF_POINTS,
        );
  const fontSize =
    source.fontSize === undefined
      ? undefined
      : normalizeEquationFontSize(source.fontSize);
  const fontSizeComplexScript =
    source.fontSizeComplexScript === undefined
      ? undefined
      : normalizeEquationFontSize(source.fontSizeComplexScript);
  if (
    fonts === null ||
    color === null ||
    underline === null ||
    border === null ||
    shading === null ||
    fitText === null ||
    verticalAlignment === null ||
    emphasisMark === null ||
    languages === null ||
    eastAsianLayout === null ||
    glow === null ||
    shadowEffect === null ||
    reflectionEffect === null ||
    textOutlineEffect === null ||
    textFillEffect === null ||
    scene3D === null ||
    properties3D === null ||
    ligatures === null ||
    numberForm === null ||
    numberSpacing === null ||
    stylisticSets === null ||
    characterSpacingTwips === null ||
    characterScalePercent === null ||
    kerningThresholdHalfPoints === null ||
    positionHalfPoints === null ||
    fontSize === null ||
    fontSizeComplexScript === null
  ) {
    return null;
  }
  if (
    source.highlight !== undefined &&
    !WORD_HIGHLIGHT_COLORS.has(
      source.highlight as WorkDocumentEquationWordHighlight,
    )
  ) {
    return null;
  }
  if (
    source.textEffect !== undefined &&
    !WORD_TEXT_EFFECTS.has(
      source.textEffect as WorkDocumentEquationWordTextEffect,
    )
  ) {
    return null;
  }
  for (const key of [
    'bold',
    'boldComplexScript',
    'italic',
    'italicComplexScript',
    'allCaps',
    'smallCaps',
    'strike',
    'doubleStrike',
    'outline',
    'shadow',
    'emboss',
    'imprint',
    'noProof',
    'snapToGrid',
    'hidden',
    'webHidden',
    'rightToLeft',
    'complexScript',
    'paragraphMarkAlwaysHidden',
  ] as const) {
    if (source[key] !== undefined && typeof source[key] !== 'boolean') {
      return null;
    }
  }
  if (equationWordRunEffectsConflict(source)) return null;
  const normalized: WorkDocumentEquationWordRunProperties = {
    ...(fonts ? { fonts } : {}),
    ...(source.bold !== undefined ? { bold: source.bold as boolean } : {}),
    ...(source.boldComplexScript !== undefined
      ? { boldComplexScript: source.boldComplexScript as boolean }
      : {}),
    ...(source.italic !== undefined
      ? { italic: source.italic as boolean }
      : {}),
    ...(source.italicComplexScript !== undefined
      ? { italicComplexScript: source.italicComplexScript as boolean }
      : {}),
    ...(source.allCaps !== undefined
      ? { allCaps: source.allCaps as boolean }
      : {}),
    ...(source.smallCaps !== undefined
      ? { smallCaps: source.smallCaps as boolean }
      : {}),
    ...(source.strike !== undefined
      ? { strike: source.strike as boolean }
      : {}),
    ...(source.doubleStrike !== undefined
      ? { doubleStrike: source.doubleStrike as boolean }
      : {}),
    ...(source.outline !== undefined
      ? { outline: source.outline as boolean }
      : {}),
    ...(source.shadow !== undefined
      ? { shadow: source.shadow as boolean }
      : {}),
    ...(source.emboss !== undefined
      ? { emboss: source.emboss as boolean }
      : {}),
    ...(source.imprint !== undefined
      ? { imprint: source.imprint as boolean }
      : {}),
    ...(source.noProof !== undefined
      ? { noProof: source.noProof as boolean }
      : {}),
    ...(source.snapToGrid !== undefined
      ? { snapToGrid: source.snapToGrid as boolean }
      : {}),
    ...(source.hidden !== undefined
      ? { hidden: source.hidden as boolean }
      : {}),
    ...(source.webHidden !== undefined
      ? { webHidden: source.webHidden as boolean }
      : {}),
    ...(color ? { color } : {}),
    ...(characterSpacingTwips !== undefined ? { characterSpacingTwips } : {}),
    ...(characterScalePercent !== undefined ? { characterScalePercent } : {}),
    ...(kerningThresholdHalfPoints !== undefined
      ? { kerningThresholdHalfPoints }
      : {}),
    ...(positionHalfPoints !== undefined ? { positionHalfPoints } : {}),
    ...(fontSize !== undefined ? { fontSize } : {}),
    ...(fontSizeComplexScript !== undefined ? { fontSizeComplexScript } : {}),
    ...(source.highlight !== undefined
      ? {
          highlight: source.highlight as WorkDocumentEquationWordHighlight,
        }
      : {}),
    ...(underline ? { underline } : {}),
    ...(source.textEffect !== undefined
      ? {
          textEffect: source.textEffect as WorkDocumentEquationWordTextEffect,
        }
      : {}),
    ...(border ? { border } : {}),
    ...(shading ? { shading } : {}),
    ...(fitText ? { fitText } : {}),
    ...(verticalAlignment ? { verticalAlignment } : {}),
    ...(source.rightToLeft !== undefined
      ? { rightToLeft: source.rightToLeft as boolean }
      : {}),
    ...(source.complexScript !== undefined
      ? { complexScript: source.complexScript as boolean }
      : {}),
    ...(emphasisMark ? { emphasisMark } : {}),
    ...(languages ? { languages } : {}),
    ...(eastAsianLayout ? { eastAsianLayout } : {}),
    ...(source.paragraphMarkAlwaysHidden !== undefined
      ? {
          paragraphMarkAlwaysHidden:
            source.paragraphMarkAlwaysHidden as boolean,
        }
      : {}),
    ...(glow ? { glow } : {}),
    ...(shadowEffect ? { shadowEffect } : {}),
    ...(reflectionEffect ? { reflectionEffect } : {}),
    ...(textOutlineEffect ? { textOutlineEffect } : {}),
    ...(textFillEffect ? { textFillEffect } : {}),
    ...(scene3D ? { scene3D } : {}),
    ...(properties3D ? { properties3D } : {}),
    ...(ligatures !== undefined ? { ligatures } : {}),
    ...(numberForm !== undefined ? { numberForm } : {}),
    ...(numberSpacing !== undefined ? { numberSpacing } : {}),
    ...(stylisticSets !== undefined ? { stylisticSets } : {}),
  };
  return Object.keys(normalized).length ? normalized : undefined;
}

function normalizeEquationWordStylisticSets(source: unknown): number[] | null {
  if (
    !Array.isArray(source) ||
    source.length > MAX_EQUATION_WORD_STYLISTIC_SET_ENTRIES
  ) {
    return null;
  }
  const normalized: number[] = [];
  const seen = new Set<number>();
  for (const id of source) {
    if (
      typeof id !== 'number' ||
      !Number.isInteger(id) ||
      id < MIN_EQUATION_WORD_STYLISTIC_SET_ID ||
      id > MAX_EQUATION_WORD_STYLISTIC_SET_ID
    ) {
      return null;
    }
    if (!seen.has(id)) {
      seen.add(id);
      normalized.push(id);
    }
  }
  return normalized;
}

function equationWordRunEffectsConflict(
  source: Record<string, unknown>,
): boolean {
  if (source.allCaps === true && source.smallCaps === true) return true;
  if (source.strike === true && source.doubleStrike === true) return true;
  if (
    source.emboss === true &&
    (source.outline === true ||
      source.shadow === true ||
      source.imprint === true)
  ) {
    return true;
  }
  return (
    source.imprint === true &&
    (source.outline === true ||
      source.shadow === true ||
      source.emboss === true)
  );
}

function normalizeEquationWordRunFonts(
  source: unknown,
): WorkDocumentEquationWordRunFonts | null | undefined {
  if (!isRecordWithKeys(source, WORD_RUN_FONT_KEYS)) return null;
  const normalized: WorkDocumentEquationWordRunFonts = {};
  for (const key of [
    'ascii',
    'highAnsi',
    'eastAsia',
    'complexScript',
  ] as const) {
    if (source[key] === undefined) continue;
    const value = normalizedEquationWordString(
      source[key],
      MAX_EQUATION_WORD_FONT_LENGTH,
    );
    if (!value) return null;
    normalized[key] = value;
  }
  for (const key of [
    'asciiTheme',
    'highAnsiTheme',
    'eastAsiaTheme',
    'complexScriptTheme',
  ] as const) {
    if (source[key] === undefined) continue;
    if (!THEME_FONTS.has(source[key] as WorkDocumentEquationThemeFont)) {
      return null;
    }
    normalized[key] = source[key] as WorkDocumentEquationThemeFont;
  }
  if (source.hint !== undefined) {
    if (
      typeof source.hint !== 'string' ||
      !['default', 'eastAsia', 'cs'].includes(source.hint)
    ) {
      return null;
    }
    normalized.hint = source.hint as WorkDocumentEquationWordRunFonts['hint'];
  }
  return Object.keys(normalized).length ? normalized : undefined;
}

function normalizeEquationWordColor(
  source: unknown,
): WorkDocumentEquationWordColor | null {
  if (!isRecordWithKeys(source, WORD_COLOR_KEYS)) return null;
  let value: WorkDocumentEquationWordColor['value'];
  if (source.value !== undefined) {
    if (source.value === 'auto') value = 'auto';
    else if (
      typeof source.value === 'string' &&
      /^#[0-9a-f]{6}$/iu.test(source.value)
    ) {
      value = source.value.toLowerCase();
    } else return null;
  }
  const theme =
    source.theme === undefined
      ? undefined
      : THEME_COLORS.has(source.theme as WorkDocumentEquationThemeColor)
        ? (source.theme as WorkDocumentEquationThemeColor)
        : null;
  const tint =
    source.tint === undefined ? undefined : normalizedByteHex(source.tint);
  const shade =
    source.shade === undefined ? undefined : normalizedByteHex(source.shade);
  if (
    theme === null ||
    tint === null ||
    shade === null ||
    (!value && (!theme || theme === 'none')) ||
    ((!theme || theme === 'none') && (tint || shade))
  ) {
    return null;
  }
  return {
    ...(value ? { value } : {}),
    ...(theme ? { theme } : {}),
    ...(tint ? { tint } : {}),
    ...(shade ? { shade } : {}),
  };
}

function normalizeEquationWordUnderline(
  source: unknown,
): WorkDocumentEquationWordUnderline | null {
  if (
    !isRecordWithKeys(source, WORD_UNDERLINE_KEYS) ||
    !UNDERLINE_STYLES.has(source.style as WorkDocumentEquationUnderlineStyle)
  ) {
    return null;
  }
  const color =
    source.color === undefined
      ? undefined
      : normalizeEquationWordColor(source.color);
  return color === null
    ? null
    : {
        style: source.style as WorkDocumentEquationUnderlineStyle,
        ...(color ? { color } : {}),
      };
}

function normalizeEquationWordRunBorder(
  source: unknown,
): WorkDocumentEquationWordRunBorder | null {
  if (
    !isRecordWithKeys(source, WORD_RUN_BORDER_KEYS) ||
    !WORD_LINE_BORDER_STYLES.has(
      source.style as WorkDocumentEquationWordLineBorderStyle,
    )
  ) {
    return null;
  }
  const color =
    source.color === undefined
      ? undefined
      : normalizeEquationWordColor(source.color);
  const sizeEighthPoints =
    source.sizeEighthPoints === undefined
      ? undefined
      : normalizeEquationInteger(
          source.sizeEighthPoints,
          MIN_EQUATION_WORD_LINE_BORDER_EIGHTH_POINTS,
          MAX_EQUATION_WORD_LINE_BORDER_EIGHTH_POINTS,
        );
  const spacingPoints =
    source.spacingPoints === undefined
      ? undefined
      : normalizeEquationInteger(
          source.spacingPoints,
          0,
          MAX_EQUATION_WORD_BORDER_SPACING_POINTS,
        );
  if (
    color === null ||
    sizeEighthPoints === null ||
    spacingPoints === null ||
    (source.shadow !== undefined && typeof source.shadow !== 'boolean') ||
    (source.frame !== undefined && typeof source.frame !== 'boolean')
  ) {
    return null;
  }
  return {
    style: source.style as WorkDocumentEquationWordLineBorderStyle,
    ...(color ? { color } : {}),
    ...(sizeEighthPoints !== undefined ? { sizeEighthPoints } : {}),
    ...(spacingPoints !== undefined ? { spacingPoints } : {}),
    ...(source.shadow !== undefined
      ? { shadow: source.shadow as boolean }
      : {}),
    ...(source.frame !== undefined ? { frame: source.frame as boolean } : {}),
  };
}

function normalizeEquationWordShading(
  source: unknown,
): WorkDocumentEquationWordShading | null {
  if (
    !isRecordWithKeys(source, WORD_SHADING_KEYS) ||
    !WORD_SHADING_PATTERNS.has(
      source.pattern as WorkDocumentEquationWordShadingPattern,
    )
  ) {
    return null;
  }
  const color =
    source.color === undefined
      ? undefined
      : normalizeEquationWordColor(source.color);
  const fill =
    source.fill === undefined
      ? undefined
      : normalizeEquationWordColor(source.fill);
  return color === null || fill === null
    ? null
    : {
        pattern: source.pattern as WorkDocumentEquationWordShadingPattern,
        ...(color ? { color } : {}),
        ...(fill ? { fill } : {}),
      };
}

function normalizeEquationWordFitText(
  source: unknown,
): WorkDocumentEquationWordFitText | null {
  if (!isRecordWithKeys(source, WORD_FIT_TEXT_KEYS)) return null;
  const widthTwips = normalizeEquationInteger(
    source.widthTwips,
    0,
    MAX_EQUATION_WORD_FIT_TEXT_WIDTH_TWIPS,
  );
  const id =
    source.id === undefined
      ? undefined
      : normalizeEquationInteger(
          source.id,
          MIN_EQUATION_WORD_FIT_TEXT_ID,
          MAX_EQUATION_WORD_FIT_TEXT_ID,
        );
  return widthTwips === null || id === null
    ? null
    : {
        widthTwips,
        ...(id !== undefined ? { id } : {}),
      };
}

function normalizeEquationWordLanguages(
  source: unknown,
): WorkDocumentEquationWordLanguages | null | undefined {
  if (!isRecordWithKeys(source, WORD_LANGUAGE_KEYS)) return null;
  const normalized: WorkDocumentEquationWordLanguages = {};
  for (const key of ['latin', 'eastAsia', 'bidi'] as const) {
    if (source[key] === undefined) continue;
    const value = normalizedEquationLanguage(source[key]);
    if (!value) return null;
    normalized[key] = value;
  }
  return Object.keys(normalized).length ? normalized : undefined;
}

function normalizeEquationWordEastAsianLayout(
  source: unknown,
): WorkDocumentEquationWordEastAsianLayout | null | undefined {
  if (!isRecordWithKeys(source, WORD_EAST_ASIAN_LAYOUT_KEYS)) return null;
  const id =
    source.id === undefined
      ? undefined
      : normalizeEquationInteger(
          source.id,
          MIN_EQUATION_WORD_EAST_ASIAN_LAYOUT_ID,
          MAX_EQUATION_WORD_EAST_ASIAN_LAYOUT_ID,
        );
  if (id === null) return null;
  for (const key of ['combine', 'vertical', 'verticalCompress'] as const) {
    if (source[key] !== undefined && typeof source[key] !== 'boolean') {
      return null;
    }
  }
  if (
    source.combineBrackets !== undefined &&
    !WORD_COMBINE_BRACKETS.has(
      source.combineBrackets as WorkDocumentEquationWordCombineBrackets,
    )
  ) {
    return null;
  }
  const normalized: WorkDocumentEquationWordEastAsianLayout = {
    ...(id !== undefined ? { id: Object.is(id, -0) ? 0 : id } : {}),
    ...(source.combine !== undefined
      ? { combine: source.combine as boolean }
      : {}),
    ...(source.combineBrackets !== undefined
      ? {
          combineBrackets:
            source.combineBrackets as WorkDocumentEquationWordCombineBrackets,
        }
      : {}),
    ...(source.vertical !== undefined
      ? { vertical: source.vertical as boolean }
      : {}),
    ...(source.verticalCompress !== undefined
      ? { verticalCompress: source.verticalCompress as boolean }
      : {}),
  };
  return Object.keys(normalized).length ? normalized : undefined;
}

function normalizeEquationWordGlow(
  source: unknown,
): WorkDocumentEquationWordGlow | null {
  if (!isRecordWithKeys(source, WORD_GLOW_KEYS)) return null;
  const radiusEmus =
    source.radiusEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.radiusEmus,
          0,
          MAX_EQUATION_WORD_GLOW_RADIUS_EMUS,
        );
  const color = normalizeEquationWordEffectColor(source.color);
  if (radiusEmus === null || color === null) return null;
  return {
    ...(radiusEmus !== undefined
      ? { radiusEmus: Object.is(radiusEmus, -0) ? 0 : radiusEmus }
      : {}),
    color,
  };
}

function normalizeEquationWordShadowEffect(
  source: unknown,
): WorkDocumentEquationWordShadowEffect | null {
  if (!isRecordWithKeys(source, WORD_SHADOW_EFFECT_KEYS)) return null;
  const blurRadiusEmus =
    source.blurRadiusEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.blurRadiusEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const distanceEmus =
    source.distanceEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.distanceEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const directionDegrees =
    source.directionDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.directionDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          0,
          MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
        );
  const horizontalScalePercent =
    source.horizontalScalePercent === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.horizontalScalePercent,
          EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
          MIN_EQUATION_WORD_EFFECT_SCALE_UNITS,
          MAX_EQUATION_WORD_EFFECT_SCALE_UNITS,
        );
  const verticalScalePercent =
    source.verticalScalePercent === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.verticalScalePercent,
          EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
          MIN_EQUATION_WORD_EFFECT_SCALE_UNITS,
          MAX_EQUATION_WORD_EFFECT_SCALE_UNITS,
        );
  const horizontalSkewDegrees =
    source.horizontalSkewDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.horizontalSkewDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          MIN_EQUATION_WORD_EFFECT_SKEW_UNITS,
          MAX_EQUATION_WORD_EFFECT_SKEW_UNITS,
        );
  const verticalSkewDegrees =
    source.verticalSkewDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.verticalSkewDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          MIN_EQUATION_WORD_EFFECT_SKEW_UNITS,
          MAX_EQUATION_WORD_EFFECT_SKEW_UNITS,
        );
  const alignment =
    source.alignment === undefined
      ? undefined
      : WORD_RECTANGLE_ALIGNMENTS.has(
            source.alignment as WorkDocumentEquationWordRectangleAlignment,
          )
        ? (source.alignment as WorkDocumentEquationWordRectangleAlignment)
        : null;
  const color = normalizeEquationWordEffectColor(source.color);
  if (
    blurRadiusEmus === null ||
    distanceEmus === null ||
    directionDegrees === null ||
    horizontalScalePercent === null ||
    verticalScalePercent === null ||
    horizontalSkewDegrees === null ||
    verticalSkewDegrees === null ||
    alignment === null ||
    color === null
  ) {
    return null;
  }
  return {
    ...(blurRadiusEmus !== undefined
      ? { blurRadiusEmus: Object.is(blurRadiusEmus, -0) ? 0 : blurRadiusEmus }
      : {}),
    ...(distanceEmus !== undefined
      ? { distanceEmus: Object.is(distanceEmus, -0) ? 0 : distanceEmus }
      : {}),
    ...(directionDegrees !== undefined ? { directionDegrees } : {}),
    ...(horizontalScalePercent !== undefined ? { horizontalScalePercent } : {}),
    ...(verticalScalePercent !== undefined ? { verticalScalePercent } : {}),
    ...(horizontalSkewDegrees !== undefined ? { horizontalSkewDegrees } : {}),
    ...(verticalSkewDegrees !== undefined ? { verticalSkewDegrees } : {}),
    ...(alignment ? { alignment } : {}),
    color,
  };
}

function normalizeEquationWordReflectionEffect(
  source: unknown,
): WorkDocumentEquationWordReflectionEffect | null {
  if (!isRecordWithKeys(source, WORD_REFLECTION_EFFECT_KEYS)) return null;
  const blurRadiusEmus =
    source.blurRadiusEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.blurRadiusEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const fixedPercentage = (value: unknown): number | null =>
    normalizeEquationScaledInteger(
      value,
      EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
      0,
      MAX_EQUATION_WORD_FIXED_PERCENTAGE_UNITS,
    );
  const startOpacityPercent =
    source.startOpacityPercent === undefined
      ? undefined
      : fixedPercentage(source.startOpacityPercent);
  const startPositionPercent =
    source.startPositionPercent === undefined
      ? undefined
      : fixedPercentage(source.startPositionPercent);
  const endOpacityPercent =
    source.endOpacityPercent === undefined
      ? undefined
      : fixedPercentage(source.endOpacityPercent);
  const endPositionPercent =
    source.endPositionPercent === undefined
      ? undefined
      : fixedPercentage(source.endPositionPercent);
  const distanceEmus =
    source.distanceEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.distanceEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const directionDegrees =
    source.directionDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.directionDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          0,
          MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
        );
  const fadeDirectionDegrees =
    source.fadeDirectionDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.fadeDirectionDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          0,
          MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
        );
  const horizontalScalePercent =
    source.horizontalScalePercent === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.horizontalScalePercent,
          EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
          MIN_EQUATION_WORD_EFFECT_SCALE_UNITS,
          MAX_EQUATION_WORD_EFFECT_SCALE_UNITS,
        );
  const verticalScalePercent =
    source.verticalScalePercent === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.verticalScalePercent,
          EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
          MIN_EQUATION_WORD_EFFECT_SCALE_UNITS,
          MAX_EQUATION_WORD_EFFECT_SCALE_UNITS,
        );
  const horizontalSkewDegrees =
    source.horizontalSkewDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.horizontalSkewDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          MIN_EQUATION_WORD_EFFECT_SKEW_UNITS,
          MAX_EQUATION_WORD_EFFECT_SKEW_UNITS,
        );
  const verticalSkewDegrees =
    source.verticalSkewDegrees === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.verticalSkewDegrees,
          EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
          MIN_EQUATION_WORD_EFFECT_SKEW_UNITS,
          MAX_EQUATION_WORD_EFFECT_SKEW_UNITS,
        );
  const alignment =
    source.alignment === undefined
      ? undefined
      : WORD_RECTANGLE_ALIGNMENTS.has(
            source.alignment as WorkDocumentEquationWordRectangleAlignment,
          )
        ? (source.alignment as WorkDocumentEquationWordRectangleAlignment)
        : null;
  if (
    blurRadiusEmus === null ||
    startOpacityPercent === null ||
    startPositionPercent === null ||
    endOpacityPercent === null ||
    endPositionPercent === null ||
    distanceEmus === null ||
    directionDegrees === null ||
    fadeDirectionDegrees === null ||
    horizontalScalePercent === null ||
    verticalScalePercent === null ||
    horizontalSkewDegrees === null ||
    verticalSkewDegrees === null ||
    alignment === null
  ) {
    return null;
  }
  return {
    ...(blurRadiusEmus !== undefined
      ? { blurRadiusEmus: Object.is(blurRadiusEmus, -0) ? 0 : blurRadiusEmus }
      : {}),
    ...(startOpacityPercent !== undefined ? { startOpacityPercent } : {}),
    ...(startPositionPercent !== undefined ? { startPositionPercent } : {}),
    ...(endOpacityPercent !== undefined ? { endOpacityPercent } : {}),
    ...(endPositionPercent !== undefined ? { endPositionPercent } : {}),
    ...(distanceEmus !== undefined
      ? { distanceEmus: Object.is(distanceEmus, -0) ? 0 : distanceEmus }
      : {}),
    ...(directionDegrees !== undefined ? { directionDegrees } : {}),
    ...(fadeDirectionDegrees !== undefined ? { fadeDirectionDegrees } : {}),
    ...(horizontalScalePercent !== undefined ? { horizontalScalePercent } : {}),
    ...(verticalScalePercent !== undefined ? { verticalScalePercent } : {}),
    ...(horizontalSkewDegrees !== undefined ? { horizontalSkewDegrees } : {}),
    ...(verticalSkewDegrees !== undefined ? { verticalSkewDegrees } : {}),
    ...(alignment ? { alignment } : {}),
  };
}

function normalizeEquationWordTextOutlineEffect(
  source: unknown,
): WorkDocumentEquationWordTextOutlineEffect | null {
  if (!isRecordWithKeys(source, WORD_TEXT_OUTLINE_EFFECT_KEYS)) return null;
  const widthEmus =
    source.widthEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.widthEmus,
          0,
          MAX_EQUATION_WORD_TEXT_OUTLINE_WIDTH_EMUS,
        );
  const cap =
    source.cap === undefined
      ? undefined
      : WORD_TEXT_OUTLINE_CAPS.has(
            source.cap as WorkDocumentEquationWordTextOutlineCap,
          )
        ? (source.cap as WorkDocumentEquationWordTextOutlineCap)
        : null;
  const compound =
    source.compound === undefined
      ? undefined
      : WORD_TEXT_OUTLINE_COMPOUNDS.has(
            source.compound as WorkDocumentEquationWordTextOutlineCompound,
          )
        ? (source.compound as WorkDocumentEquationWordTextOutlineCompound)
        : null;
  const alignment =
    source.alignment === undefined
      ? undefined
      : WORD_TEXT_OUTLINE_ALIGNMENTS.has(
            source.alignment as WorkDocumentEquationWordTextOutlineAlignment,
          )
        ? (source.alignment as WorkDocumentEquationWordTextOutlineAlignment)
        : null;
  const fill =
    source.fill === undefined
      ? undefined
      : normalizeEquationWordEffectFill(source.fill);
  const dash =
    source.dash === undefined
      ? undefined
      : normalizeEquationWordLineDash(source.dash);
  const join =
    source.join === undefined
      ? undefined
      : normalizeEquationWordLineJoin(source.join);
  if (
    widthEmus === null ||
    cap === null ||
    compound === null ||
    alignment === null ||
    fill === null ||
    dash === null ||
    join === null
  ) {
    return null;
  }
  return {
    ...(widthEmus !== undefined
      ? { widthEmus: Object.is(widthEmus, -0) ? 0 : widthEmus }
      : {}),
    ...(cap ? { cap } : {}),
    ...(compound ? { compound } : {}),
    ...(alignment ? { alignment } : {}),
    ...(fill ? { fill } : {}),
    ...(dash ? { dash } : {}),
    ...(join ? { join } : {}),
  };
}

function normalizeEquationWordTextFillEffect(
  source: unknown,
): WorkDocumentEquationWordTextFillEffect | null {
  if (!isRecordWithKeys(source, WORD_TEXT_FILL_EFFECT_KEYS)) return null;
  const fill =
    source.fill === undefined
      ? undefined
      : normalizeEquationWordEffectFill(source.fill);
  return fill === null ? null : { ...(fill ? { fill } : {}) };
}

function normalizeEquationWordScene3D(
  source: unknown,
): WorkDocumentEquationWordScene3D | null {
  if (!isRecordWithKeys(source, WORD_SCENE_3D_KEYS)) return null;
  const cameraPreset = WORD_SCENE_3D_CAMERA_PRESETS.has(
    source.cameraPreset as WorkDocumentEquationWordPresetCamera,
  )
    ? (source.cameraPreset as WorkDocumentEquationWordPresetCamera)
    : null;
  const lightRig = normalizeEquationWordScene3DLightRig(source.lightRig);
  return cameraPreset && lightRig ? { cameraPreset, lightRig } : null;
}

function normalizeEquationWordScene3DLightRig(
  source: unknown,
): WorkDocumentEquationWordScene3DLightRig | null {
  if (!isRecordWithKeys(source, WORD_SCENE_3D_LIGHT_RIG_KEYS)) return null;
  const preset = WORD_SCENE_3D_LIGHT_RIG_PRESETS.has(
    source.preset as WorkDocumentEquationWordLightRigPreset,
  )
    ? (source.preset as WorkDocumentEquationWordLightRigPreset)
    : null;
  const direction = WORD_SCENE_3D_LIGHT_RIG_DIRECTIONS.has(
    source.direction as WorkDocumentEquationWordLightRigDirection,
  )
    ? (source.direction as WorkDocumentEquationWordLightRigDirection)
    : null;
  const rotation =
    source.rotation === undefined
      ? undefined
      : normalizeEquationWordScene3DRotation(source.rotation);
  return preset && direction && rotation !== null
    ? { preset, direction, ...(rotation ? { rotation } : {}) }
    : null;
}

function normalizeEquationWordScene3DRotation(
  source: unknown,
): WorkDocumentEquationWordScene3DRotation | null {
  if (!isRecordWithKeys(source, WORD_SCENE_3D_ROTATION_KEYS)) return null;
  const latitudeDegrees = normalizeEquationScaledInteger(
    source.latitudeDegrees,
    EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
    0,
    MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
  );
  const longitudeDegrees = normalizeEquationScaledInteger(
    source.longitudeDegrees,
    EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
    0,
    MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
  );
  const revolutionDegrees = normalizeEquationScaledInteger(
    source.revolutionDegrees,
    EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
    0,
    MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
  );
  return latitudeDegrees !== null &&
    longitudeDegrees !== null &&
    revolutionDegrees !== null
    ? { latitudeDegrees, longitudeDegrees, revolutionDegrees }
    : null;
}

function normalizeEquationWordProperties3D(
  source: unknown,
): WorkDocumentEquationWordProperties3D | null {
  if (!isRecordWithKeys(source, WORD_PROPERTIES_3D_KEYS)) return null;
  const extrusionHeightEmus =
    source.extrusionHeightEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.extrusionHeightEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const contourWidthEmus =
    source.contourWidthEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.contourWidthEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const materialPreset =
    source.materialPreset === undefined
      ? undefined
      : WORD_PROPERTIES_3D_MATERIAL_PRESETS.has(
            source.materialPreset as WorkDocumentEquationWordPresetMaterial,
          )
        ? (source.materialPreset as WorkDocumentEquationWordPresetMaterial)
        : null;
  const topBevel =
    source.topBevel === undefined
      ? undefined
      : normalizeEquationWordBevel(source.topBevel);
  const bottomBevel =
    source.bottomBevel === undefined
      ? undefined
      : normalizeEquationWordBevel(source.bottomBevel);
  const extrusionColor =
    source.extrusionColor === undefined
      ? undefined
      : normalizeEquationWordEffectColor(source.extrusionColor);
  const contourColor =
    source.contourColor === undefined
      ? undefined
      : normalizeEquationWordEffectColor(source.contourColor);
  if (
    extrusionHeightEmus === null ||
    contourWidthEmus === null ||
    materialPreset === null ||
    topBevel === null ||
    bottomBevel === null ||
    extrusionColor === null ||
    contourColor === null
  ) {
    return null;
  }
  return {
    ...(extrusionHeightEmus !== undefined
      ? {
          extrusionHeightEmus: Object.is(extrusionHeightEmus, -0)
            ? 0
            : extrusionHeightEmus,
        }
      : {}),
    ...(contourWidthEmus !== undefined
      ? {
          contourWidthEmus: Object.is(contourWidthEmus, -0)
            ? 0
            : contourWidthEmus,
        }
      : {}),
    ...(materialPreset !== undefined ? { materialPreset } : {}),
    ...(topBevel !== undefined ? { topBevel } : {}),
    ...(bottomBevel !== undefined ? { bottomBevel } : {}),
    ...(extrusionColor !== undefined ? { extrusionColor } : {}),
    ...(contourColor !== undefined ? { contourColor } : {}),
  };
}

function normalizeEquationWordBevel(
  source: unknown,
): WorkDocumentEquationWordBevel | null {
  if (!isRecordWithKeys(source, WORD_PROPERTIES_3D_BEVEL_KEYS)) return null;
  const widthEmus =
    source.widthEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.widthEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const heightEmus =
    source.heightEmus === undefined
      ? undefined
      : normalizeEquationInteger(
          source.heightEmus,
          0,
          MAX_EQUATION_WORD_EFFECT_COORDINATE_EMUS,
        );
  const preset =
    source.preset === undefined
      ? undefined
      : WORD_PROPERTIES_3D_BEVEL_PRESETS.has(
            source.preset as WorkDocumentEquationWordBevelPreset,
          )
        ? (source.preset as WorkDocumentEquationWordBevelPreset)
        : null;
  if (widthEmus === null || heightEmus === null || preset === null) return null;
  return {
    ...(widthEmus !== undefined
      ? { widthEmus: Object.is(widthEmus, -0) ? 0 : widthEmus }
      : {}),
    ...(heightEmus !== undefined
      ? { heightEmus: Object.is(heightEmus, -0) ? 0 : heightEmus }
      : {}),
    ...(preset !== undefined ? { preset } : {}),
  };
}

function normalizeEquationWordEffectFill(
  source: unknown,
): WorkDocumentEquationWordEffectFill | null {
  if (!isRecord(source)) return null;
  if (source.type === 'none') {
    return isRecordWithKeys(source, WORD_EFFECT_NO_FILL_KEYS)
      ? { type: 'none' }
      : null;
  }
  if (source.type === 'solid') {
    if (!isRecordWithKeys(source, WORD_EFFECT_SOLID_FILL_KEYS)) return null;
    const color =
      source.color === undefined
        ? undefined
        : normalizeEquationWordEffectColor(source.color);
    return color === null
      ? null
      : { type: 'solid', ...(color ? { color } : {}) };
  }
  if (
    source.type !== 'gradient' ||
    !isRecordWithKeys(source, WORD_EFFECT_GRADIENT_FILL_KEYS)
  ) {
    return null;
  }
  const stops = normalizeEquationWordGradientStops(source.stops);
  const shade =
    source.shade === undefined
      ? undefined
      : normalizeEquationWordGradientShade(source.shade);
  return stops === null || shade === null
    ? null
    : {
        type: 'gradient',
        ...(stops ? { stops } : {}),
        ...(shade ? { shade } : {}),
      };
}

function normalizeEquationWordGradientStops(
  source: unknown,
): WorkDocumentEquationWordGradientStop[] | null | undefined {
  if (source === undefined) return undefined;
  if (
    !Array.isArray(source) ||
    source.length < MIN_EQUATION_WORD_GRADIENT_STOPS ||
    source.length > MAX_EQUATION_WORD_GRADIENT_STOPS
  ) {
    return null;
  }
  const stops: WorkDocumentEquationWordGradientStop[] = [];
  for (const stop of source) {
    if (!isRecordWithKeys(stop, WORD_GRADIENT_STOP_KEYS)) return null;
    const positionPercent = normalizeEquationScaledInteger(
      stop.positionPercent,
      EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
      0,
      MAX_EQUATION_WORD_FIXED_PERCENTAGE_UNITS,
    );
    const color = normalizeEquationWordEffectColor(stop.color);
    if (positionPercent === null || color === null) return null;
    stops.push({ positionPercent, color });
  }
  return stops;
}

function normalizeEquationWordGradientShade(
  source: unknown,
): WorkDocumentEquationWordGradientShade | null {
  if (!isRecord(source)) return null;
  if (source.type === 'linear') {
    if (!isRecordWithKeys(source, WORD_LINEAR_GRADIENT_SHADE_KEYS)) return null;
    const angleDegrees =
      source.angleDegrees === undefined
        ? undefined
        : normalizeEquationScaledInteger(
            source.angleDegrees,
            EQUATION_WORD_ANGLE_UNITS_PER_DEGREE,
            0,
            MAX_EQUATION_WORD_EFFECT_DIRECTION_UNITS,
          );
    if (
      angleDegrees === null ||
      (source.scaled !== undefined && typeof source.scaled !== 'boolean')
    ) {
      return null;
    }
    return {
      type: 'linear',
      ...(angleDegrees !== undefined ? { angleDegrees } : {}),
      ...(source.scaled !== undefined
        ? { scaled: source.scaled as boolean }
        : {}),
    };
  }
  if (
    source.type !== 'path' ||
    !isRecordWithKeys(source, WORD_PATH_GRADIENT_SHADE_KEYS)
  ) {
    return null;
  }
  const path =
    source.path === undefined
      ? undefined
      : WORD_GRADIENT_PATHS.has(
            source.path as WorkDocumentEquationWordGradientPath,
          )
        ? (source.path as WorkDocumentEquationWordGradientPath)
        : null;
  const fillToRectangle =
    source.fillToRectangle === undefined
      ? undefined
      : normalizeEquationWordGradientFillRectangle(source.fillToRectangle);
  return path === null || fillToRectangle === null
    ? null
    : {
        type: 'path',
        ...(path ? { path } : {}),
        ...(fillToRectangle ? { fillToRectangle } : {}),
      };
}

function normalizeEquationWordGradientFillRectangle(
  source: unknown,
): WorkDocumentEquationWordGradientFillRectangle | null {
  if (!isRecordWithKeys(source, WORD_GRADIENT_FILL_RECTANGLE_KEYS)) return null;
  const percentage = (value: unknown): number | null =>
    normalizeEquationScaledInteger(
      value,
      EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
      MIN_EQUATION_WORD_EFFECT_SCALE_UNITS,
      MAX_EQUATION_WORD_EFFECT_SCALE_UNITS,
    );
  const leftPercent =
    source.leftPercent === undefined
      ? undefined
      : percentage(source.leftPercent);
  const topPercent =
    source.topPercent === undefined ? undefined : percentage(source.topPercent);
  const rightPercent =
    source.rightPercent === undefined
      ? undefined
      : percentage(source.rightPercent);
  const bottomPercent =
    source.bottomPercent === undefined
      ? undefined
      : percentage(source.bottomPercent);
  if (
    leftPercent === null ||
    topPercent === null ||
    rightPercent === null ||
    bottomPercent === null
  ) {
    return null;
  }
  return {
    ...(leftPercent !== undefined ? { leftPercent } : {}),
    ...(topPercent !== undefined ? { topPercent } : {}),
    ...(rightPercent !== undefined ? { rightPercent } : {}),
    ...(bottomPercent !== undefined ? { bottomPercent } : {}),
  };
}

function normalizeEquationWordLineDash(
  source: unknown,
): WorkDocumentEquationWordLineDash | null {
  if (!isRecordWithKeys(source, WORD_LINE_DASH_KEYS)) return null;
  if (source.preset === undefined) return {};
  return WORD_PRESET_LINE_DASHES.has(
    source.preset as WorkDocumentEquationWordPresetLineDash,
  )
    ? { preset: source.preset as WorkDocumentEquationWordPresetLineDash }
    : null;
}

function normalizeEquationWordLineJoin(
  source: unknown,
): WorkDocumentEquationWordLineJoin | null {
  if (!isRecord(source)) return null;
  if (source.type === 'round' || source.type === 'bevel') {
    return isRecordWithKeys(source, WORD_ROUND_OR_BEVEL_LINE_JOIN_KEYS)
      ? { type: source.type }
      : null;
  }
  if (
    source.type !== 'miter' ||
    !isRecordWithKeys(source, WORD_MITER_LINE_JOIN_KEYS)
  ) {
    return null;
  }
  const limitPercent =
    source.limitPercent === undefined
      ? undefined
      : normalizeEquationScaledInteger(
          source.limitPercent,
          EQUATION_WORD_PERCENTAGE_UNITS_PER_PERCENT,
          0,
          MAX_EQUATION_WORD_EFFECT_SCALE_UNITS,
        );
  return limitPercent === null
    ? null
    : {
        type: 'miter',
        ...(limitPercent !== undefined ? { limitPercent } : {}),
      };
}

function normalizeEquationWordEffectColor(
  source: unknown,
): WorkDocumentEquationWordEffectColor | null {
  if (
    !isRecordWithKeys(source, WORD_EFFECT_COLOR_KEYS) ||
    (source.type !== 'rgb' && source.type !== 'scheme')
  ) {
    return null;
  }
  const transforms = normalizeEquationWordColorTransforms(source.transforms);
  if (transforms === null) return null;
  if (source.type === 'rgb') {
    if (
      typeof source.value !== 'string' ||
      !/^#[0-9a-f]{6}$/iu.test(source.value)
    ) {
      return null;
    }
    return {
      type: 'rgb',
      value: source.value.toLowerCase(),
      ...(transforms?.length ? { transforms } : {}),
    };
  }
  if (
    !WORD_EFFECT_SCHEME_COLORS.has(
      source.value as WorkDocumentEquationWordEffectSchemeColor,
    )
  ) {
    return null;
  }
  return {
    type: 'scheme',
    value: source.value as WorkDocumentEquationWordEffectSchemeColor,
    ...(transforms?.length ? { transforms } : {}),
  };
}

function normalizeEquationWordColorTransforms(
  source: unknown,
): WorkDocumentEquationWordColorTransform[] | null | undefined {
  if (source === undefined) return undefined;
  if (
    !Array.isArray(source) ||
    source.length > MAX_EQUATION_WORD_COLOR_TRANSFORMS
  ) {
    return null;
  }
  const normalized: WorkDocumentEquationWordColorTransform[] = [];
  for (const transform of source) {
    if (!isRecordWithKeys(transform, WORD_COLOR_TRANSFORM_KEYS)) return null;
    const type = transform.type as WorkDocumentEquationWordColorTransformType;
    if (
      !WORD_COLOR_TRANSFORM_TYPES.has(type) ||
      !Number.isInteger(transform.value)
    ) {
      return null;
    }
    const minimum = WORD_FIXED_COLOR_TRANSFORM_TYPES.has(type)
      ? 0
      : type === 'hueMod'
        ? 0
        : MIN_EQUATION_WORD_COLOR_PERCENTAGE;
    const maximum = WORD_FIXED_COLOR_TRANSFORM_TYPES.has(type)
      ? MAX_EQUATION_WORD_FIXED_COLOR_PERCENTAGE
      : MAX_EQUATION_WORD_COLOR_PERCENTAGE;
    const value = transform.value as number;
    if (value < minimum || value > maximum) return null;
    normalized.push({ type, value: Object.is(value, -0) ? 0 : value });
  }
  return normalized.length ? normalized : undefined;
}

function normalizeEquationFontSize(source: unknown): number | null {
  return typeof source === 'number' &&
    Number.isFinite(source) &&
    source > 0 &&
    source <= MAX_EQUATION_FONT_SIZE &&
    Number.isInteger(source * 2)
    ? source
    : null;
}

function normalizeEquationInteger(
  source: unknown,
  minimum: number,
  maximum: number,
): number | null {
  return typeof source === 'number' &&
    Number.isInteger(source) &&
    source >= minimum &&
    source <= maximum
    ? source
    : null;
}

function normalizeEquationScaledInteger(
  source: unknown,
  scale: number,
  minimum: number,
  maximum: number,
): number | null {
  if (typeof source !== 'number' || !Number.isFinite(source)) return null;
  const scaled = source * scale;
  if (!Number.isSafeInteger(scaled) || scaled < minimum || scaled > maximum) {
    return null;
  }
  const normalized = scaled / scale;
  return Object.is(normalized, -0) ? 0 : normalized;
}

function normalizedEquationWordString(
  source: unknown,
  maximumLength: number,
): string | null {
  if (typeof source !== 'string') return null;
  const normalized = source.trim();
  return normalized &&
    normalized.length <= maximumLength &&
    !/[\p{Cc}\p{Cs}]/u.test(normalized) &&
    validXmlText(normalized)
    ? normalized
    : null;
}

function normalizedEquationLanguage(source: unknown): string | null {
  const normalized = normalizedEquationWordString(
    source,
    MAX_EQUATION_LANGUAGE_LENGTH,
  );
  return normalized &&
    /^(?:x-none|[a-z0-9]{1,8}(?:-[a-z0-9]{1,8})*)$/iu.test(normalized)
    ? normalized
    : null;
}

function normalizedByteHex(source: unknown): string | null {
  return typeof source === 'string' && /^[0-9a-f]{2}$/iu.test(source.trim())
    ? source.trim().toUpperCase()
    : null;
}

function isRecordWithKeys(
  source: unknown,
  allowed: ReadonlySet<string>,
): source is Record<string, unknown> {
  return (
    isRecord(source) && Object.keys(source).every((key) => allowed.has(key))
  );
}

function boundedInteger(
  source: unknown,
  minimum: number,
  maximum: number,
): source is number {
  return (
    typeof source === 'number' &&
    Number.isInteger(source) &&
    source >= minimum &&
    source <= maximum
  );
}

function normalizeManualBreak(
  source: unknown,
): WorkDocumentEquationManualBreak | null {
  if (!isRecord(source)) return null;
  if (source.alignmentAt === undefined) return {};
  return Number.isInteger(source.alignmentAt) &&
    Number(source.alignmentAt) >= 1 &&
    Number(source.alignmentAt) <= 255
    ? { alignmentAt: Number(source.alignmentAt) }
    : null;
}

function runMathVariant(
  expression: Extract<WorkDocumentEquationExpression, { type: 'run' }>,
): string | null {
  if (!expression.script && !expression.style && !expression.normalText) {
    return null;
  }
  const script = expression.script ?? 'roman';
  const style =
    expression.style ?? (expression.normalText ? 'plain' : 'italic');
  if (script === 'doubleStruck') return 'double-struck';
  if (script === 'monospace') return 'monospace';
  if (script === 'fraktur') {
    return style === 'bold' || style === 'boldItalic'
      ? 'bold-fraktur'
      : 'fraktur';
  }
  if (script === 'script') {
    return style === 'bold' || style === 'boldItalic'
      ? 'bold-script'
      : 'script';
  }
  if (script === 'sansSerif') {
    if (style === 'plain') return 'sans-serif';
    if (style === 'bold') return 'bold-sans-serif';
    if (style === 'boldItalic') return 'sans-serif-bold-italic';
    return 'sans-serif-italic';
  }
  if (style === 'plain') return 'normal';
  if (style === 'bold') return 'bold';
  if (style === 'boldItalic') return 'bold-italic';
  return 'italic';
}

function wordRunMathMlAttributes(
  expression: Extract<WorkDocumentEquationExpression, { type: 'run' }>,
): Record<string, string> {
  return wordPropertiesMathMlAttributes(
    expression.wordRunProperties,
    expression.text,
  );
}

function controlMathMlAttributes(
  expression: Exclude<WorkDocumentEquationExpression, { type: 'run' }>,
  text: string,
): Record<string, string> {
  return wordPropertiesMathMlAttributes(expression.controlProperties, text);
}

function wordPropertiesMathMlAttributes(
  properties: WorkDocumentEquationWordRunProperties | undefined,
  text: string,
): Record<string, string> {
  if (!properties) return {};
  const complexScript = wordPropertiesUseComplexScript(properties, text);
  const eastAsia = !complexScript && wordTextUsesEastAsianScript(text);
  const font = complexScript
    ? properties.fonts?.complexScript
    : eastAsia
      ? properties.fonts?.eastAsia
      : (properties.fonts?.ascii ?? properties.fonts?.highAnsi);
  const size = complexScript
    ? (properties.fontSizeComplexScript ?? properties.fontSize)
    : properties.fontSize;
  const language = complexScript
    ? properties.languages?.bidi
    : eastAsia
      ? properties.languages?.eastAsia
      : properties.languages?.latin;
  const bold = complexScript
    ? (properties.boldComplexScript ?? properties.bold)
    : properties.bold;
  const italic = complexScript
    ? (properties.italicComplexScript ?? properties.italic)
    : properties.italic;
  const styles = [
    font ? `font-family:${cssString(font)}` : '',
    bold === undefined ? '' : `font-weight:${bold ? 'bold' : 'normal'}`,
    italic === undefined ? '' : `font-style:${italic ? 'italic' : 'normal'}`,
    wordRunCaseStyles(properties),
    wordRunLigatureStyles(properties.ligatures),
    wordRunNumberStyles(properties.numberForm, properties.numberSpacing),
    wordRunStylisticSetStyles(properties.stylisticSets),
    wordRunTextDecoration(properties),
    wordRunBorderStyles(properties),
    properties.characterSpacingTwips === undefined
      ? ''
      : `letter-spacing:${properties.characterSpacingTwips / 20}pt`,
    properties.characterScalePercent === undefined
      ? ''
      : `font-stretch:${properties.characterScalePercent}%`,
    wordRunMathMlKerning(properties, size),
    properties.positionHalfPoints === undefined
      ? ''
      : `vertical-align:${properties.positionHalfPoints / 2}pt`,
    wordRunVerticalAlignmentStyles(properties.verticalAlignment),
    wordRunEmphasisMarkStyles(properties.emphasisMark),
  ].filter(Boolean);
  const color =
    wordRunTextFillMathMlColor(properties.textFillEffect) ??
    properties.color?.value;
  const background = wordRunMathMlBackground(properties);
  return {
    ...(color && color !== 'auto' ? { mathcolor: color } : {}),
    ...(background ? { mathbackground: background } : {}),
    ...(size !== undefined ? { mathsize: `${size}pt` } : {}),
    ...(properties.rightToLeft !== undefined
      ? { dir: properties.rightToLeft ? 'rtl' : 'ltr' }
      : {}),
    ...(language ? { lang: language } : {}),
    ...(styles.length ? { style: styles.join(';') } : {}),
  };
}

function wordRunLigatureStyles(
  ligatures: WorkDocumentEquationWordLigatures | undefined,
): string {
  if (ligatures === undefined) return '';
  const flags = WORD_LIGATURE_FLAGS.get(ligatures);
  if (flags === undefined) return '';
  if (flags === 0) return 'font-variant-ligatures:none';
  return `font-variant-ligatures:${[
    flags & WORD_LIGATURE_STANDARD ? 'common-ligatures' : 'no-common-ligatures',
    flags & WORD_LIGATURE_CONTEXTUAL ? 'contextual' : 'no-contextual',
    flags & WORD_LIGATURE_HISTORICAL
      ? 'historical-ligatures'
      : 'no-historical-ligatures',
    flags & WORD_LIGATURE_DISCRETIONAL
      ? 'discretionary-ligatures'
      : 'no-discretionary-ligatures',
  ].join(' ')}`;
}

function wordRunNumberStyles(
  numberForm: WorkDocumentEquationWordNumberForm | undefined,
  numberSpacing: WorkDocumentEquationWordNumberSpacing | undefined,
): string {
  if (numberForm === undefined && numberSpacing === undefined) return '';
  const values = [
    numberForm === 'lining'
      ? 'lining-nums'
      : numberForm === 'oldStyle'
        ? 'oldstyle-nums'
        : '',
    numberSpacing === 'proportional'
      ? 'proportional-nums'
      : numberSpacing === 'tabular'
        ? 'tabular-nums'
        : '',
  ].filter(Boolean);
  return `font-variant-numeric:${values.length ? values.join(' ') : 'normal'}`;
}

function wordRunStylisticSetStyles(
  stylisticSets: number[] | undefined,
): string {
  if (stylisticSets === undefined) return '';
  if (!stylisticSets.length) return 'font-feature-settings:normal';
  return `font-feature-settings:${stylisticSets
    .map((id) => `"ss${String(id).padStart(2, '0')}" 1`)
    .join(', ')}`;
}

function wordRunTextFillMathMlColor(
  effect: WorkDocumentEquationWordTextFillEffect | undefined,
): string | undefined {
  if (!effect) return undefined;
  const fill = effect.fill;
  if (!fill) return '#000000';
  if (fill.type === 'gradient') {
    return fill.stops ? undefined : '#000000';
  }
  if (fill.type !== 'solid') return undefined;
  if (!fill.color) return '#000000';
  return fill.color.type === 'rgb' && !fill.color.transforms?.length
    ? fill.color.value
    : undefined;
}

function wordRunVerticalAlignmentStyles(
  alignment: WorkDocumentEquationWordVerticalAlignment | undefined,
): string {
  if (!alignment) return '';
  if (alignment === 'baseline') return 'vertical-align:baseline';
  return `vertical-align:${alignment === 'superscript' ? 'super' : 'sub'};font-size:smaller`;
}

function wordRunEmphasisMarkStyles(
  emphasisMark: WorkDocumentEquationWordEmphasisMark | undefined,
): string {
  if (!emphasisMark) return '';
  if (emphasisMark === 'none') return 'text-emphasis-style:none';
  const style =
    emphasisMark === 'comma'
      ? cssString(',')
      : emphasisMark === 'circle'
        ? 'open circle'
        : 'filled dot';
  const position = emphasisMark === 'underDot' ? 'under right' : 'over right';
  return `text-emphasis-style:${style};text-emphasis-position:${position}`;
}

function wordRunCaseStyles(
  properties: WorkDocumentEquationWordRunProperties,
): string {
  return [
    properties.allCaps === undefined
      ? ''
      : `text-transform:${properties.allCaps ? 'uppercase' : 'none'}`,
    properties.smallCaps === undefined
      ? ''
      : `font-variant-caps:${properties.smallCaps ? 'small-caps' : 'normal'}`,
  ]
    .filter(Boolean)
    .join(';');
}

function wordRunMathMlKerning(
  properties: WorkDocumentEquationWordRunProperties,
  fontSize: number | undefined,
): string {
  const threshold = properties.kerningThresholdHalfPoints;
  if (threshold === undefined) return '';
  if (threshold === 0) return 'font-kerning:normal';
  if (fontSize === undefined) return '';
  return `font-kerning:${fontSize * 2 >= threshold ? 'normal' : 'none'}`;
}

function wordRunMathMlBackground(
  properties: WorkDocumentEquationWordRunProperties,
): string | undefined {
  if (properties.highlight) {
    return properties.highlight === 'none'
      ? 'transparent'
      : WORD_HIGHLIGHT_MATHML_COLORS[properties.highlight];
  }
  const shading = properties.shading;
  if (!shading) return undefined;
  if (shading.pattern === 'nil') return 'transparent';
  if (shading.pattern === 'clear') {
    if (!shading.fill || shading.fill.value === 'auto') return 'transparent';
    return shading.fill.value;
  }
  if (shading.pattern === 'solid') {
    const color = shading.color?.value;
    return color && color !== 'auto' ? color : undefined;
  }
  return undefined;
}

function wordPropertiesUseComplexScript(
  properties: WorkDocumentEquationWordRunProperties,
  text: string,
): boolean {
  return (
    properties.complexScript === true ||
    properties.rightToLeft === true ||
    /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}]/u.test(
      text,
    )
  );
}

function wordTextUsesEastAsianScript(text: string): boolean {
  return /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}]/u.test(
    text,
  );
}

function wordRunTextDecoration(
  properties: WorkDocumentEquationWordRunProperties,
): string {
  const underline = properties.underline;
  const lines = [
    underline && underline.style !== 'none' ? 'underline' : '',
    properties.strike || properties.doubleStrike ? 'line-through' : '',
  ].filter(Boolean);
  if (!lines.length) return '';
  const underlineStyle = underline
    ? equationUnderlineCssStyle(underline.style)
    : 'solid';
  const style =
    properties.doubleStrike || underlineStyle === 'double'
      ? 'double'
      : underlineStyle;
  const color = underline?.color?.value;
  return [
    `text-decoration-line:${lines.join(' ')}`,
    `text-decoration-style:${style}`,
    color && color !== 'auto' ? `text-decoration-color:${color}` : '',
  ]
    .filter(Boolean)
    .join(';');
}

function wordRunBorderStyles(
  properties: WorkDocumentEquationWordRunProperties,
): string {
  const border = properties.border;
  if (!border) return '';
  if (border.style === 'nil' || border.style === 'none') {
    return 'border-style:none';
  }
  const style = equationWordLineBorderCssStyle(border.style);
  if (!style || border.sizeEighthPoints === undefined) return '';
  const directColor = border.color?.value;
  if (border.color && directColor === undefined) return '';
  const color =
    directColor && directColor !== 'auto' ? directColor : 'currentColor';
  return [
    `border:${border.sizeEighthPoints / 8}pt ${style} ${color}`,
    border.spacingPoints === undefined
      ? ''
      : `padding:${border.spacingPoints}pt`,
  ]
    .filter(Boolean)
    .join(';');
}

function equationWordLineBorderCssStyle(
  style: WorkDocumentEquationWordLineBorderStyle,
): 'solid' | 'double' | 'dotted' | 'dashed' | 'outset' | 'inset' | null {
  if (style === 'single' || style === 'thick') return 'solid';
  if (style === 'double') return 'double';
  if (style === 'dotted') return 'dotted';
  if (style === 'dashed' || style === 'dashSmallGap') return 'dashed';
  if (style === 'outset' || style === 'inset') return style;
  return null;
}

function equationUnderlineCssStyle(
  style: WorkDocumentEquationUnderlineStyle,
): 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy' {
  if (style === 'double' || style === 'wavyDouble') return 'double';
  if (style === 'dotted' || style === 'dottedHeavy') return 'dotted';
  if (/dash/iu.test(style)) return 'dashed';
  if (/wav/iu.test(style)) return 'wavy';
  return 'solid';
}

function cssString(source: string): string {
  return `"${source.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`;
}

function borderBoxNotation(
  expression: Extract<WorkDocumentEquationExpression, { type: 'borderBox' }>,
): string {
  const edges = [
    expression.hideTop ? '' : 'top',
    expression.hideBottom ? '' : 'bottom',
    expression.hideLeft ? '' : 'left',
    expression.hideRight ? '' : 'right',
  ].filter(Boolean);
  const notations = edges.length === 4 ? ['box'] : edges;
  if (expression.strikeHorizontal) notations.push('horizontalstrike');
  if (expression.strikeVertical) notations.push('verticalstrike');
  if (expression.strikeBottomLeftToTopRight) {
    notations.push('updiagonalstrike');
  }
  if (expression.strikeTopLeftToBottomRight) {
    notations.push('downdiagonalstrike');
  }
  return notations.join(' ') || 'none';
}

function validXmlText(source: string): boolean {
  for (const character of Array.from(source)) {
    const codePoint = character.codePointAt(0) ?? 0;
    if (
      (codePoint < 0x20 &&
        codePoint !== 0x09 &&
        codePoint !== 0x0a &&
        codePoint !== 0x0d) ||
      (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
      (codePoint & 0xffff) === 0xfffe ||
      (codePoint & 0xffff) === 0xffff
    ) {
      return false;
    }
  }
  return true;
}

function isRecord(source: unknown): source is Record<string, unknown> {
  return (
    Boolean(source) && typeof source === 'object' && !Array.isArray(source)
  );
}
