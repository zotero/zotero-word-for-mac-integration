/*
 * Word.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class WordApplication, WordDocument, WordWindow, WordCommandBarControl, WordCommandBarButton, WordCommandBarCombobox, WordCommandBarPopup, WordCommandBar, WordDocumentProperty, WordCustomDocumentProperty, WordWebPageFont, WordWordComment, WordWordList, WordWordOptions, WordAddIn, WordApplication, WordAutoTextEntry, WordBookmark, WordBorderOptions, WordBorder, WordBrowser, WordBuildingBlockCategory, WordBuildingBlockType, WordBuildingBlock, WordCaptionLabel, WordCheckBox, WordCoauthLock, WordCoauthUpdate, WordCoauthor, WordCoauthoring, WordConflict, WordCustomLabel, WordDataMergeDataField, WordDataMergeDataSource, WordDataMergeFieldName, WordDataMergeField, WordDataMerge, WordDefaultWebOptions, WordDialog, WordDocumentVersion, WordDocument, WordDropCap, WordDropDown, WordEndnoteOptions, WordEndnote, WordEnvelope, WordFieldOptions, WordField, WordFileConverter, WordFind, WordFont, WordFootnoteOptions, WordFootnote, WordFormField, WordFrame, WordHeaderFooter, WordHeadingStyle, WordHyperlinkObject, WordIndex, WordKeyBinding, WordLetterContent, WordLineNumbering, WordLinkFormat, WordListEntry, WordListFormat, WordListGallery, WordListLevel, WordListTemplate, WordMailingLabel, WordMathAccent, WordMathAutocorrectEntry, WordMathAutocorrect, WordMathBar, WordMathBorderBox, WordMathBox, WordMathBreak, WordMathDelimiter, WordMathEquationArray, WordMathFraction, WordMathFunc, WordMathFunction, WordMathGroupChar, WordMathLeftScripts, WordMathLowerLimit, WordMathMatrixColumn, WordMathMatrixRow, WordMathMatrix, WordMathNary, WordMathObject, WordMathPhantom, WordMathRadical, WordMathRecognizedFunction, WordMathSubAndSuperScript, WordMathSubscript, WordMathSuperscript, WordMathUpperLimit, WordPageNumberOptions, WordPageNumber, WordPageSetup, WordPane, WordRangeEndnoteOptions, WordRangeFootnoteOptions, WordRecentFile, WordReplacement, WordReviewer, WordRevision, WordSelectionObject, WordSubdocument, WordSystemObject, WordTabStop, WordTableOfAuthorities, WordTableOfContents, WordTableOfFigures, WordTemplate, WordTextColumn, WordTextInput, WordTextRetrievalMode, WordVariable, WordView, WordWebOptions, WordWindow, WordZoom, WordAdjustment, WordCalloutFormat, WordFillFormat, WordGlowFormat, WordHorizontalLineFormat, WordInlineShape, WordInlineHorizontalLine, WordInlinePictureBullet, WordInlinePicture, WordLineFormat, WordOfficeTheme, WordPictureFormat, WordReflectionFormat, WordShadowFormat, WordShape, WordCallout, WordLineShape, WordPicture, WordSoftEdgeFormat, WordStandardInlineHorizontalLine, WordTextBox, WordTextFrame, WordThemeColorScheme, WordThemeColor, WordThemeEffectScheme, WordThemeFontScheme, WordThemeFont, WordMajorThemeFont, WordMinorThemeFont, WordThemeFonts, WordThreeDFormat, WordWordArtFormat, WordWordArt, WordWrapFormat, WordWordStyle, WordParagraphFormat, WordParagraph, WordSection, WordShading, WordTextRange, WordCharacter, WordGrammaticalError, WordSentence, WordSpellingError, WordStoryRange, WordWord, WordAutocorrectEntry, WordAutocorrect, WordDictionary, WordFirstLetterException, WordLanguage, WordOtherCorrectionsException, WordReadabilityStatistic, WordSynonymInfo, WordTwoInitialCapsException, WordCell, WordColumnOptions, WordColumn, WordCondition, WordRowOptions, WordRow, WordTableStyle, WordTable;

enum WordSaveOptions {
	WordSaveOptionsYes = 'yes ' /* Save the file. */,
	WordSaveOptionsNo = 'no  ' /* Do not save the file. */,
	WordSaveOptionsAsk = 'ask ' /* Ask the user whether or not to save the file. */
};
typedef enum WordSaveOptions WordSaveOptions;

enum WordPrintingErrorHandling {
	WordPrintingErrorHandlingStandard = 'lwst' /* Standard PostScript error handling */,
	WordPrintingErrorHandlingDetailed = 'lwdt' /* print a detailed report of PostScript errors */
};
typedef enum WordPrintingErrorHandling WordPrintingErrorHandling;

enum WordMsoLineDashStyle {
	WordMsoLineDashStyleLineDashStyleUnset = '\000\222\377\376',
	WordMsoLineDashStyleLineDashStyleSolid = '\000\223\000\001',
	WordMsoLineDashStyleLineDashStyleSquareDot = '\000\223\000\002',
	WordMsoLineDashStyleLineDashStyleRoundDot = '\000\223\000\003',
	WordMsoLineDashStyleLineDashStyleDash = '\000\223\000\004',
	WordMsoLineDashStyleLineDashStyleDashDot = '\000\223\000\005',
	WordMsoLineDashStyleLineDashStyleDashDotDot = '\000\223\000\006',
	WordMsoLineDashStyleLineDashStyleLongDash = '\000\223\000\007',
	WordMsoLineDashStyleLineDashStyleLongDashDot = '\000\223\000\010',
	WordMsoLineDashStyleLineDashStyleLongDashDotDot = '\000\223\000\011',
	WordMsoLineDashStyleLineDashStyleSystemDash = '\000\223\000\012',
	WordMsoLineDashStyleLineDashStyleSystemDot = '\000\223\000\013',
	WordMsoLineDashStyleLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum WordMsoLineDashStyle WordMsoLineDashStyle;

enum WordMsoLineStyle {
	WordMsoLineStyleLineStyleUnset = '\000\224\377\376',
	WordMsoLineStyleSingleLine = '\000\225\000\001',
	WordMsoLineStyleThinThinLine = '\000\225\000\002',
	WordMsoLineStyleThinThickLine = '\000\225\000\003',
	WordMsoLineStyleThickThinLine = '\000\225\000\004',
	WordMsoLineStyleThickBetweenThinLine = '\000\225\000\005'
};
typedef enum WordMsoLineStyle WordMsoLineStyle;

enum WordMsoArrowheadStyle {
	WordMsoArrowheadStyleArrowheadStyleUnset = '\000\221\377\376',
	WordMsoArrowheadStyleNoArrowhead = '\000\222\000\001',
	WordMsoArrowheadStyleTriangleArrowhead = '\000\222\000\002',
	WordMsoArrowheadStyleOpen_arrowhead = '\000\222\000\003',
	WordMsoArrowheadStyleStealthArrowhead = '\000\222\000\004',
	WordMsoArrowheadStyleDiamondArrowhead = '\000\222\000\005',
	WordMsoArrowheadStyleOvalArrowhead = '\000\222\000\006'
};
typedef enum WordMsoArrowheadStyle WordMsoArrowheadStyle;

enum WordMsoArrowheadWidth {
	WordMsoArrowheadWidthArrowheadWidthUnset = '\000\220\377\376',
	WordMsoArrowheadWidthNarrowWidthArrowhead = '\000\221\000\001',
	WordMsoArrowheadWidthMediumWidthArrowhead = '\000\221\000\002',
	WordMsoArrowheadWidthWideArrowhead = '\000\221\000\003'
};
typedef enum WordMsoArrowheadWidth WordMsoArrowheadWidth;

enum WordMsoArrowheadLength {
	WordMsoArrowheadLengthArrowheadLengthUnset = '\000\223\377\376',
	WordMsoArrowheadLengthShortArrowhead = '\000\224\000\001',
	WordMsoArrowheadLengthMediumArrowhead = '\000\224\000\002',
	WordMsoArrowheadLengthLongArrowhead = '\000\224\000\003'
};
typedef enum WordMsoArrowheadLength WordMsoArrowheadLength;

enum WordMsoFillType {
	WordMsoFillTypeFillUnset = '\000c\377\376',
	WordMsoFillTypeFillSolid = '\000d\000\001',
	WordMsoFillTypeFillPatterned = '\000d\000\002',
	WordMsoFillTypeFillGradient = '\000d\000\003',
	WordMsoFillTypeFillTextured = '\000d\000\004',
	WordMsoFillTypeFillBackground = '\000d\000\005',
	WordMsoFillTypeFillPicture = '\000d\000\006'
};
typedef enum WordMsoFillType WordMsoFillType;

enum WordMsoGradientStyle {
	WordMsoGradientStyleGradientUnset = '\000d\377\376',
	WordMsoGradientStyleHorizontalGradient = '\000e\000\001',
	WordMsoGradientStyleVerticalGradient = '\000e\000\002',
	WordMsoGradientStyleDiagonalUpGradient = '\000e\000\003',
	WordMsoGradientStyleDiagonalDownGradient = '\000e\000\004',
	WordMsoGradientStyleFromCornerGradient = '\000e\000\005',
	WordMsoGradientStyleFromTitleGradient = '\000e\000\006',
	WordMsoGradientStyleFromCenterGradient = '\000e\000\007'
};
typedef enum WordMsoGradientStyle WordMsoGradientStyle;

enum WordMsoGradientColorType {
	WordMsoGradientColorTypeGradientTypeUnset = '\003\357\377\376',
	WordMsoGradientColorTypeSingleShadeGradientType = '\003\360\000\001',
	WordMsoGradientColorTypeTwoColorsGradientType = '\003\360\000\002',
	WordMsoGradientColorTypePresetColorsGradientType = '\003\360\000\003',
	WordMsoGradientColorTypeMultiColorsGradientType = '\003\360\000\004'
};
typedef enum WordMsoGradientColorType WordMsoGradientColorType;

enum WordMsoTextureType {
	WordMsoTextureTypeTextureTypeTextureTypeUnset = '\003\360\377\376',
	WordMsoTextureTypeTextureTypePresetTexture = '\003\361\000\001',
	WordMsoTextureTypeTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum WordMsoTextureType WordMsoTextureType;

enum WordMsoPresetTexture {
	WordMsoPresetTexturePresetTextureUnset = '\000e\377\376',
	WordMsoPresetTextureTexturePapyrus = '\000f\000\001',
	WordMsoPresetTextureTextureCanvas = '\000f\000\002',
	WordMsoPresetTextureTextureDenim = '\000f\000\003',
	WordMsoPresetTextureTextureWovenMat = '\000f\000\004',
	WordMsoPresetTextureTextureWaterDroplets = '\000f\000\005',
	WordMsoPresetTextureTexturePaperBag = '\000f\000\006',
	WordMsoPresetTextureTextureFishFossil = '\000f\000\007',
	WordMsoPresetTextureTextureSand = '\000f\000\010',
	WordMsoPresetTextureTextureGreenMarble = '\000f\000\011',
	WordMsoPresetTextureTextureWhiteMarble = '\000f\000\012',
	WordMsoPresetTextureTextureBrownMarble = '\000f\000\013',
	WordMsoPresetTextureTextureGranite = '\000f\000\014',
	WordMsoPresetTextureTextureNewsprint = '\000f\000\015',
	WordMsoPresetTextureTextureRecycledPaper = '\000f\000\016',
	WordMsoPresetTextureTextureParchment = '\000f\000\017',
	WordMsoPresetTextureTextureStationery = '\000f\000\020',
	WordMsoPresetTextureTextureBlueTissuePaper = '\000f\000\021',
	WordMsoPresetTextureTexturePinkTissuePaper = '\000f\000\022',
	WordMsoPresetTextureTexturePurpleMesh = '\000f\000\023',
	WordMsoPresetTextureTextureBouquet = '\000f\000\024',
	WordMsoPresetTextureTextureCork = '\000f\000\025',
	WordMsoPresetTextureTextureWalnut = '\000f\000\026',
	WordMsoPresetTextureTextureOak = '\000f\000\027',
	WordMsoPresetTextureTextureMediumWood = '\000f\000\030'
};
typedef enum WordMsoPresetTexture WordMsoPresetTexture;

enum WordMsoPatternType {
	WordMsoPatternTypePatternUnset = '\000f\377\376',
	WordMsoPatternTypeFivePercentPattern = '\000g\000\001',
	WordMsoPatternTypeTenPercentPattern = '\000g\000\002',
	WordMsoPatternTypeTwentyPercentPattern = '\000g\000\003',
	WordMsoPatternTypeTwentyFivePercentPattern = '\000g\000\004',
	WordMsoPatternTypeThirtyPercentPattern = '\000g\000\005',
	WordMsoPatternTypeFortyPercentPattern = '\000g\000\006',
	WordMsoPatternTypeFiftyPercentPattern = '\000g\000\007',
	WordMsoPatternTypeSixtyPercentPattern = '\000g\000\010',
	WordMsoPatternTypeSeventyPercentPattern = '\000g\000\011',
	WordMsoPatternTypeSeventyFivePercentPattern = '\000g\000\012',
	WordMsoPatternTypeEightyPercentPattern = '\000g\000\013',
	WordMsoPatternTypeNinetyPercentPattern = '\000g\000\014',
	WordMsoPatternTypeDarkHorizontalPattern = '\000g\000\015',
	WordMsoPatternTypeDarkVerticalPattern = '\000g\000\016',
	WordMsoPatternTypeDarkDownwardDiagonalPattern = '\000g\000\017',
	WordMsoPatternTypeDarkUpwardDiagonalPattern = '\000g\000\020',
	WordMsoPatternTypeSmallCheckerBoardPattern = '\000g\000\021',
	WordMsoPatternTypeTrellisPattern = '\000g\000\022',
	WordMsoPatternTypeLightHorizontalPattern = '\000g\000\023',
	WordMsoPatternTypeLightVerticalPattern = '\000g\000\024',
	WordMsoPatternTypeLightDownwardDiagonalPattern = '\000g\000\025',
	WordMsoPatternTypeLightUpwardDiagonalPattern = '\000g\000\026',
	WordMsoPatternTypeSmallGridPattern = '\000g\000\027',
	WordMsoPatternTypeDottedDiamondPattern = '\000g\000\030',
	WordMsoPatternTypeWideDownwardDiagonal = '\000g\000\031',
	WordMsoPatternTypeWideUpwardDiagonalPattern = '\000g\000\032',
	WordMsoPatternTypeDashedUpwardDiagonalPattern = '\000g\000\033',
	WordMsoPatternTypeDashedDownwardDiagonalPattern = '\000g\000\034',
	WordMsoPatternTypeNarrowVerticalPattern = '\000g\000\035',
	WordMsoPatternTypeNarrowHorizontalPattern = '\000g\000\036',
	WordMsoPatternTypeDashedVerticalPattern = '\000g\000\037',
	WordMsoPatternTypeDashedHorizontalPattern = '\000g\000 ',
	WordMsoPatternTypeLargeConfettiPattern = '\000g\000!',
	WordMsoPatternTypeLargeGridPattern = '\000g\000\"',
	WordMsoPatternTypeHorizontalBrickPattern = '\000g\000#',
	WordMsoPatternTypeLargeCheckerBoardPattern = '\000g\000$',
	WordMsoPatternTypeSmallConfettiPattern = '\000g\000%',
	WordMsoPatternTypeZigZagPattern = '\000g\000&',
	WordMsoPatternTypeSolidDiamondPattern = '\000g\000\'',
	WordMsoPatternTypeDiagonalBrickPattern = '\000g\000(',
	WordMsoPatternTypeOutlinedDiamondPattern = '\000g\000)',
	WordMsoPatternTypePlaidPattern = '\000g\000*',
	WordMsoPatternTypeSpherePattern = '\000g\000+',
	WordMsoPatternTypeWeavePattern = '\000g\000,',
	WordMsoPatternTypeDottedGridPattern = '\000g\000-',
	WordMsoPatternTypeDivotPattern = '\000g\000.',
	WordMsoPatternTypeShinglePattern = '\000g\000/',
	WordMsoPatternTypeWavePattern = '\000g\0000',
	WordMsoPatternTypeHorizontalPattern = '\000g\0001',
	WordMsoPatternTypeVerticalPattern = '\000g\0002',
	WordMsoPatternTypeCrossPattern = '\000g\0003',
	WordMsoPatternTypeDownwardDiagonalPattern = '\000g\0004',
	WordMsoPatternTypeUpwardDiagonalPattern = '\000g\0005',
	WordMsoPatternTypeDiagonalCrossPattern = '\000g\0005'
};
typedef enum WordMsoPatternType WordMsoPatternType;

enum WordMsoPresetGradientType {
	WordMsoPresetGradientTypePresetGradientUnset = '\000g\377\376',
	WordMsoPresetGradientTypeGradientEarlySunset = '\000h\000\001',
	WordMsoPresetGradientTypeGradientLateSunset = '\000h\000\002',
	WordMsoPresetGradientTypeGradientNightfall = '\000h\000\003',
	WordMsoPresetGradientTypeGradientDaybreak = '\000h\000\004',
	WordMsoPresetGradientTypeGradientHorizon = '\000h\000\005',
	WordMsoPresetGradientTypeGradientDesert = '\000h\000\006',
	WordMsoPresetGradientTypeGradientOcean = '\000h\000\007',
	WordMsoPresetGradientTypeGradientCalmWater = '\000h\000\010',
	WordMsoPresetGradientTypeGradientFire = '\000h\000\011',
	WordMsoPresetGradientTypeGradientFog = '\000h\000\012',
	WordMsoPresetGradientTypeGradientMoss = '\000h\000\013',
	WordMsoPresetGradientTypeGradientPeacock = '\000h\000\014',
	WordMsoPresetGradientTypeGradientWheat = '\000h\000\015',
	WordMsoPresetGradientTypeGradientParchment = '\000h\000\016',
	WordMsoPresetGradientTypeGradientMahogany = '\000h\000\017',
	WordMsoPresetGradientTypeGradientRainbow = '\000h\000\020',
	WordMsoPresetGradientTypeGradientRainbow2 = '\000h\000\021',
	WordMsoPresetGradientTypeGradientGold = '\000h\000\022',
	WordMsoPresetGradientTypeGradientGold2 = '\000h\000\023',
	WordMsoPresetGradientTypeGradientBrass = '\000h\000\024',
	WordMsoPresetGradientTypeGradientChrome = '\000h\000\025',
	WordMsoPresetGradientTypeGradientChrome2 = '\000h\000\026',
	WordMsoPresetGradientTypeGradientSilver = '\000h\000\027',
	WordMsoPresetGradientTypeGradientSapphire = '\000h\000\030'
};
typedef enum WordMsoPresetGradientType WordMsoPresetGradientType;

enum WordMsoShadowType {
	WordMsoShadowTypeShadowUnset = '\003_\377\376',
	WordMsoShadowTypeShadow1 = '\003`\000\001',
	WordMsoShadowTypeShadow2 = '\003`\000\002',
	WordMsoShadowTypeShadow3 = '\003`\000\003',
	WordMsoShadowTypeShadow4 = '\003`\000\004',
	WordMsoShadowTypeShadow5 = '\003`\000\005',
	WordMsoShadowTypeShadow6 = '\003`\000\006',
	WordMsoShadowTypeShadow7 = '\003`\000\007',
	WordMsoShadowTypeShadow8 = '\003`\000\010',
	WordMsoShadowTypeShadow9 = '\003`\000\011',
	WordMsoShadowTypeShadow10 = '\003`\000\012',
	WordMsoShadowTypeShadow11 = '\003`\000\013',
	WordMsoShadowTypeShadow12 = '\003`\000\014',
	WordMsoShadowTypeShadow13 = '\003`\000\015',
	WordMsoShadowTypeShadow14 = '\003`\000\016',
	WordMsoShadowTypeShadow15 = '\003`\000\017',
	WordMsoShadowTypeShadow16 = '\003`\000\020',
	WordMsoShadowTypeShadow17 = '\003`\000\021',
	WordMsoShadowTypeShadow18 = '\003`\000\022',
	WordMsoShadowTypeShadow19 = '\003`\000\023',
	WordMsoShadowTypeShadow20 = '\003`\000\024',
	WordMsoShadowTypeShadow21 = '\003`\000\025',
	WordMsoShadowTypeShadow22 = '\003`\000\026',
	WordMsoShadowTypeShadow23 = '\003`\000\027',
	WordMsoShadowTypeShadow24 = '\003`\000\030',
	WordMsoShadowTypeShadow25 = '\003`\000\031',
	WordMsoShadowTypeShadow26 = '\003`\000\032',
	WordMsoShadowTypeShadow27 = '\003`\000\033',
	WordMsoShadowTypeShadow28 = '\003`\000\034',
	WordMsoShadowTypeShadow29 = '\003`\000\035',
	WordMsoShadowTypeShadow30 = '\003`\000\036',
	WordMsoShadowTypeShadow31 = '\003`\000\037',
	WordMsoShadowTypeShadow32 = '\003`\000 ',
	WordMsoShadowTypeShadow33 = '\003`\000!',
	WordMsoShadowTypeShadow34 = '\003`\000\"',
	WordMsoShadowTypeShadow35 = '\003`\000#',
	WordMsoShadowTypeShadow36 = '\003`\000$',
	WordMsoShadowTypeShadow37 = '\003`\000%',
	WordMsoShadowTypeShadow38 = '\003`\000&',
	WordMsoShadowTypeShadow39 = '\003`\000\'',
	WordMsoShadowTypeShadow40 = '\003`\000(',
	WordMsoShadowTypeShadow41 = '\003`\000)',
	WordMsoShadowTypeShadow42 = '\003`\000*',
	WordMsoShadowTypeShadow43 = '\003`\000+'
};
typedef enum WordMsoShadowType WordMsoShadowType;

enum WordMsoPresetTextEffect {
	WordMsoPresetTextEffectWordartFormatUnset = '\003\361\377\376',
	WordMsoPresetTextEffectWordartFormat1 = '\003\362\000\000',
	WordMsoPresetTextEffectWordartFormat2 = '\003\362\000\001',
	WordMsoPresetTextEffectWordartFormat3 = '\003\362\000\002',
	WordMsoPresetTextEffectWordartFormat4 = '\003\362\000\003',
	WordMsoPresetTextEffectWordartFormat5 = '\003\362\000\004',
	WordMsoPresetTextEffectWordartFormat6 = '\003\362\000\005',
	WordMsoPresetTextEffectWordartFormat7 = '\003\362\000\006',
	WordMsoPresetTextEffectWordartFormat8 = '\003\362\000\007',
	WordMsoPresetTextEffectWordartFormat9 = '\003\362\000\010',
	WordMsoPresetTextEffectWordartFormat10 = '\003\362\000\011',
	WordMsoPresetTextEffectWordartFormat11 = '\003\362\000\012',
	WordMsoPresetTextEffectWordartFormat12 = '\003\362\000\013',
	WordMsoPresetTextEffectWordartFormat13 = '\003\362\000\014',
	WordMsoPresetTextEffectWordartFormat14 = '\003\362\000\015',
	WordMsoPresetTextEffectWordartFormat15 = '\003\362\000\016',
	WordMsoPresetTextEffectWordartFormat16 = '\003\362\000\017',
	WordMsoPresetTextEffectWordartFormat17 = '\003\362\000\020',
	WordMsoPresetTextEffectWordartFormat18 = '\003\362\000\021',
	WordMsoPresetTextEffectWordartFormat19 = '\003\362\000\022',
	WordMsoPresetTextEffectWordartFormat20 = '\003\362\000\023',
	WordMsoPresetTextEffectWordartFormat21 = '\003\362\000\024',
	WordMsoPresetTextEffectWordartFormat22 = '\003\362\000\025',
	WordMsoPresetTextEffectWordartFormat23 = '\003\362\000\026',
	WordMsoPresetTextEffectWordartFormat24 = '\003\362\000\027',
	WordMsoPresetTextEffectWordartFormat25 = '\003\362\000\030',
	WordMsoPresetTextEffectWordartFormat26 = '\003\362\000\031',
	WordMsoPresetTextEffectWordartFormat27 = '\003\362\000\032',
	WordMsoPresetTextEffectWordartFormat28 = '\003\362\000\033',
	WordMsoPresetTextEffectWordartFormat29 = '\003\362\000\034',
	WordMsoPresetTextEffectWordartFormat30 = '\003\362\000\035',
	WordMsoPresetTextEffectWordartFormat31 = '\003\362\000\036',
	WordMsoPresetTextEffectWordartFormat32 = '\003\362\000\037',
	WordMsoPresetTextEffectWordartFormat33 = '\003\362\000 ',
	WordMsoPresetTextEffectWordartFormat34 = '\003\362\000!',
	WordMsoPresetTextEffectWordartFormat35 = '\003\362\000\"',
	WordMsoPresetTextEffectWordartFormat36 = '\003\362\000#',
	WordMsoPresetTextEffectWordartFormat37 = '\003\362\000$',
	WordMsoPresetTextEffectWordartFormat38 = '\003\362\000%',
	WordMsoPresetTextEffectWordartFormat39 = '\003\362\000&',
	WordMsoPresetTextEffectWordartFormat40 = '\003\362\000\'',
	WordMsoPresetTextEffectWordartFormat41 = '\003\362\000(',
	WordMsoPresetTextEffectWordartFormat42 = '\003\362\000)',
	WordMsoPresetTextEffectWordartFormat43 = '\003\362\000*',
	WordMsoPresetTextEffectWordartFormat44 = '\003\362\000+',
	WordMsoPresetTextEffectWordartFormat45 = '\003\362\000,',
	WordMsoPresetTextEffectWordartFormat46 = '\003\362\000-',
	WordMsoPresetTextEffectWordartFormat47 = '\003\362\000.',
	WordMsoPresetTextEffectWordartFormat48 = '\003\362\000/',
	WordMsoPresetTextEffectWordartFormat49 = '\003\362\0000',
	WordMsoPresetTextEffectWordartFormat50 = '\003\362\0001'
};
typedef enum WordMsoPresetTextEffect WordMsoPresetTextEffect;

enum WordMsoPresetTextEffectShape {
	WordMsoPresetTextEffectShapeTextEffectShapeUnset = '\000\227\377\376',
	WordMsoPresetTextEffectShapePlainText = '\000\230\000\001',
	WordMsoPresetTextEffectShapeStop = '\000\230\000\002',
	WordMsoPresetTextEffectShapeTriangleUp = '\000\230\000\003',
	WordMsoPresetTextEffectShapeTriangleDown = '\000\230\000\004',
	WordMsoPresetTextEffectShapeChevronUp = '\000\230\000\005',
	WordMsoPresetTextEffectShapeChevronDown = '\000\230\000\006',
	WordMsoPresetTextEffectShapeRingInside = '\000\230\000\007',
	WordMsoPresetTextEffectShapeRingOutside = '\000\230\000\010',
	WordMsoPresetTextEffectShapeArchUpCurve = '\000\230\000\011',
	WordMsoPresetTextEffectShapeArchDownCurve = '\000\230\000\012',
	WordMsoPresetTextEffectShapeCircleCurve = '\000\230\000\013',
	WordMsoPresetTextEffectShapeButtonCurve = '\000\230\000\014',
	WordMsoPresetTextEffectShapeArchUpPour = '\000\230\000\015',
	WordMsoPresetTextEffectShapeArchDownPour = '\000\230\000\016',
	WordMsoPresetTextEffectShapeCirclePour = '\000\230\000\017',
	WordMsoPresetTextEffectShapeButtonPour = '\000\230\000\020',
	WordMsoPresetTextEffectShapeCurveUp = '\000\230\000\021',
	WordMsoPresetTextEffectShapeCurveDown = '\000\230\000\022',
	WordMsoPresetTextEffectShapeCanUp = '\000\230\000\023',
	WordMsoPresetTextEffectShapeCanDown = '\000\230\000\024',
	WordMsoPresetTextEffectShapeWave1 = '\000\230\000\025',
	WordMsoPresetTextEffectShapeWave2 = '\000\230\000\026',
	WordMsoPresetTextEffectShapeDoubleWave1 = '\000\230\000\027',
	WordMsoPresetTextEffectShapeDoubleWave2 = '\000\230\000\030',
	WordMsoPresetTextEffectShapeInflate = '\000\230\000\031',
	WordMsoPresetTextEffectShapeDeflate = '\000\230\000\032',
	WordMsoPresetTextEffectShapeInflateBottom = '\000\230\000\033',
	WordMsoPresetTextEffectShapeDeflateBottom = '\000\230\000\034',
	WordMsoPresetTextEffectShapeInflateTop = '\000\230\000\035',
	WordMsoPresetTextEffectShapeDeflateTop = '\000\230\000\036',
	WordMsoPresetTextEffectShapeDeflateInflate = '\000\230\000\037',
	WordMsoPresetTextEffectShapeDeflateInflateDeflate = '\000\230\000 ',
	WordMsoPresetTextEffectShapeFadeRight = '\000\230\000!',
	WordMsoPresetTextEffectShapeFadeLeft = '\000\230\000\"',
	WordMsoPresetTextEffectShapeFadeUp = '\000\230\000#',
	WordMsoPresetTextEffectShapeFadeDown = '\000\230\000$',
	WordMsoPresetTextEffectShapeSlantUp = '\000\230\000%',
	WordMsoPresetTextEffectShapeSlantDown = '\000\230\000&',
	WordMsoPresetTextEffectShapeCascadeUp = '\000\230\000\'',
	WordMsoPresetTextEffectShapeCascadeDown = '\000\230\000('
};
typedef enum WordMsoPresetTextEffectShape WordMsoPresetTextEffectShape;

enum WordMsoTextEffectAlignment {
	WordMsoTextEffectAlignmentTextEffectAlignmentUnset = '\000\226\377\376',
	WordMsoTextEffectAlignmentLeftTextEffectAlignment = '\000\227\000\001',
	WordMsoTextEffectAlignmentCenteredTextEffectAlignment = '\000\227\000\002',
	WordMsoTextEffectAlignmentRightTextEffectAlignment = '\000\227\000\003',
	WordMsoTextEffectAlignmentJustifyTextEffectAlignment = '\000\227\000\004',
	WordMsoTextEffectAlignmentWordJustifyTextEffectAlignment = '\000\227\000\005',
	WordMsoTextEffectAlignmentStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum WordMsoTextEffectAlignment WordMsoTextEffectAlignment;

enum WordMsoPresetLightingDirection {
	WordMsoPresetLightingDirectionPresetLightingDirectionUnset = '\000\233\377\376',
	WordMsoPresetLightingDirectionLightFromTopLeft = '\000\234\000\001',
	WordMsoPresetLightingDirectionLightFromTop = '\000\234\000\002',
	WordMsoPresetLightingDirectionLightFromTopRight = '\000\234\000\003',
	WordMsoPresetLightingDirectionLightFromLeft = '\000\234\000\004',
	WordMsoPresetLightingDirectionLightFromNone = '\000\234\000\005',
	WordMsoPresetLightingDirectionLightFromRight = '\000\234\000\006',
	WordMsoPresetLightingDirectionLightFromBottomLeft = '\000\234\000\007',
	WordMsoPresetLightingDirectionLightFromBottom = '\000\234\000\010',
	WordMsoPresetLightingDirectionLightFromBottomRight = '\000\234\000\011'
};
typedef enum WordMsoPresetLightingDirection WordMsoPresetLightingDirection;

enum WordMsoPresetLightingSoftness {
	WordMsoPresetLightingSoftnessLightingSoftnessUnset = '\000\234\377\376',
	WordMsoPresetLightingSoftnessLightingDim = '\000\235\000\001',
	WordMsoPresetLightingSoftnessLightingNormal = '\000\235\000\002',
	WordMsoPresetLightingSoftnessLightingBright = '\000\235\000\003'
};
typedef enum WordMsoPresetLightingSoftness WordMsoPresetLightingSoftness;

enum WordMsoPresetMaterial {
	WordMsoPresetMaterialPresetMaterialUnset = '\000\235\377\376',
	WordMsoPresetMaterialMatte = '\000\236\000\001',
	WordMsoPresetMaterialPlastic = '\000\236\000\002',
	WordMsoPresetMaterialMetal = '\000\236\000\003',
	WordMsoPresetMaterialWireframe = '\000\236\000\004',
	WordMsoPresetMaterialMatte2 = '\000\236\000\005',
	WordMsoPresetMaterialPlastic2 = '\000\236\000\006',
	WordMsoPresetMaterialMetal2 = '\000\236\000\007',
	WordMsoPresetMaterialWarmMatte = '\000\236\000\010',
	WordMsoPresetMaterialTranslucentPowder = '\000\236\000\011',
	WordMsoPresetMaterialPowder = '\000\236\000\012',
	WordMsoPresetMaterialDarkEdge = '\000\236\000\013',
	WordMsoPresetMaterialSoftEdge = '\000\236\000\014',
	WordMsoPresetMaterialMaterialClear = '\000\236\000\015',
	WordMsoPresetMaterialFlat = '\000\236\000\016',
	WordMsoPresetMaterialSoftMetal = '\000\236\000\017'
};
typedef enum WordMsoPresetMaterial WordMsoPresetMaterial;

enum WordMsoPresetExtrusionDirection {
	WordMsoPresetExtrusionDirectionPresetExtrusionDirectionUnset = '\000\231\377\376',
	WordMsoPresetExtrusionDirectionExtrudeBottomRight = '\000\232\000\001',
	WordMsoPresetExtrusionDirectionExtrudeBottom = '\000\232\000\002',
	WordMsoPresetExtrusionDirectionExtrudeBottomLeft = '\000\232\000\003',
	WordMsoPresetExtrusionDirectionExtrudeRight = '\000\232\000\004',
	WordMsoPresetExtrusionDirectionExtrudeNone = '\000\232\000\005',
	WordMsoPresetExtrusionDirectionExtrudeLeft = '\000\232\000\006',
	WordMsoPresetExtrusionDirectionExtrudeTopRight = '\000\232\000\007',
	WordMsoPresetExtrusionDirectionExtrudeTop = '\000\232\000\010',
	WordMsoPresetExtrusionDirectionExtrudeTopLeft = '\000\232\000\011'
};
typedef enum WordMsoPresetExtrusionDirection WordMsoPresetExtrusionDirection;

enum WordMsoPresetThreeDFormat {
	WordMsoPresetThreeDFormatPresetThreeDFormatUnset = '\000\230\377\376',
	WordMsoPresetThreeDFormatFormat1 = '\000\231\000\001',
	WordMsoPresetThreeDFormatFormat2 = '\000\231\000\002',
	WordMsoPresetThreeDFormatFormat3 = '\000\231\000\003',
	WordMsoPresetThreeDFormatFormat4 = '\000\231\000\004',
	WordMsoPresetThreeDFormatFormat5 = '\000\231\000\005',
	WordMsoPresetThreeDFormatFormat6 = '\000\231\000\006',
	WordMsoPresetThreeDFormatFormat7 = '\000\231\000\007',
	WordMsoPresetThreeDFormatFormat8 = '\000\231\000\010',
	WordMsoPresetThreeDFormatFormat9 = '\000\231\000\011',
	WordMsoPresetThreeDFormatFormat10 = '\000\231\000\012',
	WordMsoPresetThreeDFormatFormat11 = '\000\231\000\013',
	WordMsoPresetThreeDFormatFormat12 = '\000\231\000\014',
	WordMsoPresetThreeDFormatFormat13 = '\000\231\000\015',
	WordMsoPresetThreeDFormatFormat14 = '\000\231\000\016',
	WordMsoPresetThreeDFormatFormat15 = '\000\231\000\017',
	WordMsoPresetThreeDFormatFormat16 = '\000\231\000\020',
	WordMsoPresetThreeDFormatFormat17 = '\000\231\000\021',
	WordMsoPresetThreeDFormatFormat18 = '\000\231\000\022',
	WordMsoPresetThreeDFormatFormat19 = '\000\231\000\023',
	WordMsoPresetThreeDFormatFormat20 = '\000\231\000\024'
};
typedef enum WordMsoPresetThreeDFormat WordMsoPresetThreeDFormat;

enum WordMsoExtrusionColorType {
	WordMsoExtrusionColorTypeExtrusionColorUnset = '\000\232\377\376',
	WordMsoExtrusionColorTypeExtrusionColorAutomatic = '\000\233\000\001',
	WordMsoExtrusionColorTypeExtrusionColorCustom = '\000\233\000\002'
};
typedef enum WordMsoExtrusionColorType WordMsoExtrusionColorType;

enum WordMsoConnectorType {
	WordMsoConnectorTypeConnectorTypeUnset = '\000h\377\376',
	WordMsoConnectorTypeStraight = '\000i\000\001',
	WordMsoConnectorTypeElbow = '\000i\000\002',
	WordMsoConnectorTypeCurve = '\000i\000\003'
};
typedef enum WordMsoConnectorType WordMsoConnectorType;

enum WordMsoHorizontalAnchor {
	WordMsoHorizontalAnchorHorizontalAnchorUnset = '\000\236\377\376',
	WordMsoHorizontalAnchorHorizontalAnchorNone = '\000\237\000\001',
	WordMsoHorizontalAnchorHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum WordMsoHorizontalAnchor WordMsoHorizontalAnchor;

enum WordMsoVerticalAnchor {
	WordMsoVerticalAnchorVerticalAnchorUnset = '\000\237\377\376',
	WordMsoVerticalAnchorAnchorTop = '\000\240\000\001',
	WordMsoVerticalAnchorAnchorTopBaseline = '\000\240\000\002',
	WordMsoVerticalAnchorAnchorMiddle = '\000\240\000\003',
	WordMsoVerticalAnchorAnchorBottom = '\000\240\000\004',
	WordMsoVerticalAnchorAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum WordMsoVerticalAnchor WordMsoVerticalAnchor;

enum WordMsoAutoShapeType {
	WordMsoAutoShapeTypeAutoshapeShapeTypeUnset = '\000i\377\376',
	WordMsoAutoShapeTypeAutoshapeRectangle = '\000j\000\001',
	WordMsoAutoShapeTypeAutoshapeParallelogram = '\000j\000\002',
	WordMsoAutoShapeTypeAutoshapeTrapezoid = '\000j\000\003',
	WordMsoAutoShapeTypeAutoshapeDiamond = '\000j\000\004',
	WordMsoAutoShapeTypeAutoshapeRoundedRectangle = '\000j\000\005',
	WordMsoAutoShapeTypeAutoshapeOctagon = '\000j\000\006',
	WordMsoAutoShapeTypeAutoshapeIsoscelesTriangle = '\000j\000\007',
	WordMsoAutoShapeTypeAutoshapeRightTriangle = '\000j\000\010',
	WordMsoAutoShapeTypeAutoshapeOval = '\000j\000\011',
	WordMsoAutoShapeTypeAutoshapeHexagon = '\000j\000\012',
	WordMsoAutoShapeTypeAutoshapeCross = '\000j\000\013',
	WordMsoAutoShapeTypeAutoshapeRegularPentagon = '\000j\000\014',
	WordMsoAutoShapeTypeAutoshapeCan = '\000j\000\015',
	WordMsoAutoShapeTypeAutoshapeCube = '\000j\000\016',
	WordMsoAutoShapeTypeAutoshapeBevel = '\000j\000\017',
	WordMsoAutoShapeTypeAutoshapeFoldedCorner = '\000j\000\020',
	WordMsoAutoShapeTypeAutoshapeSmileyFace = '\000j\000\021',
	WordMsoAutoShapeTypeAutoshapeDonut = '\000j\000\022',
	WordMsoAutoShapeTypeAutoshapeNoSymbol = '\000j\000\023',
	WordMsoAutoShapeTypeAutoshapeBlockArc = '\000j\000\024',
	WordMsoAutoShapeTypeAutoshapeHeart = '\000j\000\025',
	WordMsoAutoShapeTypeAutoshapeLightningBolt = '\000j\000\026',
	WordMsoAutoShapeTypeAutoshapeSun = '\000j\000\027',
	WordMsoAutoShapeTypeAutoshapeMoon = '\000j\000\030',
	WordMsoAutoShapeTypeAutoshapeArc = '\000j\000\031',
	WordMsoAutoShapeTypeAutoshapeDoubleBracket = '\000j\000\032',
	WordMsoAutoShapeTypeAutoshapeDoubleBrace = '\000j\000\033',
	WordMsoAutoShapeTypeAutoshapePlaque = '\000j\000\034',
	WordMsoAutoShapeTypeAutoshapeLeftBracket = '\000j\000\035',
	WordMsoAutoShapeTypeAutoshapeRightBracket = '\000j\000\036',
	WordMsoAutoShapeTypeAutoshapeLeftBrace = '\000j\000\037',
	WordMsoAutoShapeTypeAutoshapeRightBrace = '\000j\000 ',
	WordMsoAutoShapeTypeAutoshapeRightArrow = '\000j\000!',
	WordMsoAutoShapeTypeAutoshapeLeftArrow = '\000j\000\"',
	WordMsoAutoShapeTypeAutoshapeUpArrow = '\000j\000#',
	WordMsoAutoShapeTypeAutoshapeDownArrow = '\000j\000$',
	WordMsoAutoShapeTypeAutoshapeLeftRightArrow = '\000j\000%',
	WordMsoAutoShapeTypeAutoshapeUpDownArrow = '\000j\000&',
	WordMsoAutoShapeTypeAutoshapeQuadArrow = '\000j\000\'',
	WordMsoAutoShapeTypeAutoshapeLeftRightUpArrow = '\000j\000(',
	WordMsoAutoShapeTypeAutoshapeBentArrow = '\000j\000)',
	WordMsoAutoShapeTypeAutoshapeUTurnArrow = '\000j\000*',
	WordMsoAutoShapeTypeAutoshapeLeftUpArrow = '\000j\000+',
	WordMsoAutoShapeTypeAutoshapeBentUpArrow = '\000j\000,',
	WordMsoAutoShapeTypeAutoshapeCurvedRightArrow = '\000j\000-',
	WordMsoAutoShapeTypeAutoshapeCurvedLeftArrow = '\000j\000.',
	WordMsoAutoShapeTypeAutoshapeCurvedUpArrow = '\000j\000/',
	WordMsoAutoShapeTypeAutoshapeCurvedDownArrow = '\000j\0000',
	WordMsoAutoShapeTypeAutoshapeStripedRightArrow = '\000j\0001',
	WordMsoAutoShapeTypeAutoshapeNotchedRightArrow = '\000j\0002',
	WordMsoAutoShapeTypeAutoshapePentagon = '\000j\0003',
	WordMsoAutoShapeTypeAutoshapeChevron = '\000j\0004',
	WordMsoAutoShapeTypeAutoshapeRightArrowCallout = '\000j\0005',
	WordMsoAutoShapeTypeAutoshapeLeftArrowCallout = '\000j\0006',
	WordMsoAutoShapeTypeAutoshapeUpArrowCallout = '\000j\0007',
	WordMsoAutoShapeTypeAutoshapeDownArrowCallout = '\000j\0008',
	WordMsoAutoShapeTypeAutoshapeLeftRightArrowCallout = '\000j\0009',
	WordMsoAutoShapeTypeAutoshapeUpDownArrowCallout = '\000j\000:',
	WordMsoAutoShapeTypeAutoshapeQuadArrowCallout = '\000j\000;',
	WordMsoAutoShapeTypeAutoshapeCircularArrow = '\000j\000<',
	WordMsoAutoShapeTypeAutoshapeFlowchartProcess = '\000j\000=',
	WordMsoAutoShapeTypeAutoshapeFlowchartAlternateProcess = '\000j\000>',
	WordMsoAutoShapeTypeAutoshapeFlowchartDecision = '\000j\000\?',
	WordMsoAutoShapeTypeAutoshapeFlowchartData = '\000j\000@',
	WordMsoAutoShapeTypeAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	WordMsoAutoShapeTypeAutoshapeFlowchartInternalStorage = '\000j\000B',
	WordMsoAutoShapeTypeAutoshapeFlowchartDocument = '\000j\000C',
	WordMsoAutoShapeTypeAutoshapeFlowchartMultiDocument = '\000j\000D',
	WordMsoAutoShapeTypeAutoshapeFlowchartTerminator = '\000j\000E',
	WordMsoAutoShapeTypeAutoshapeFlowchartPreparation = '\000j\000F',
	WordMsoAutoShapeTypeAutoshapeFlowchartManualInput = '\000j\000G',
	WordMsoAutoShapeTypeAutoshapeFlowchartManualOperation = '\000j\000H',
	WordMsoAutoShapeTypeAutoshapeFlowchartConnector = '\000j\000I',
	WordMsoAutoShapeTypeAutoshapeFlowchartOffpageConnector = '\000j\000J',
	WordMsoAutoShapeTypeAutoshapeFlowchartCard = '\000j\000K',
	WordMsoAutoShapeTypeAutoshapeFlowchartPunchedTape = '\000j\000L',
	WordMsoAutoShapeTypeAutoshapeFlowchartSummingJunction = '\000j\000M',
	WordMsoAutoShapeTypeAutoshapeFlowchartOr = '\000j\000N',
	WordMsoAutoShapeTypeAutoshapeFlowchartCollate = '\000j\000O',
	WordMsoAutoShapeTypeAutoshapeFlowchartSort = '\000j\000P',
	WordMsoAutoShapeTypeAutoshapeFlowchartExtract = '\000j\000Q',
	WordMsoAutoShapeTypeAutoshapeFlowchartMerge = '\000j\000R',
	WordMsoAutoShapeTypeAutoshapeFlowchartStoredData = '\000j\000S',
	WordMsoAutoShapeTypeAutoshapeFlowchartDelay = '\000j\000T',
	WordMsoAutoShapeTypeAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	WordMsoAutoShapeTypeAutoshapeFlowchartMagneticDisk = '\000j\000V',
	WordMsoAutoShapeTypeAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	WordMsoAutoShapeTypeAutoshapeFlowchartDisplay = '\000j\000X',
	WordMsoAutoShapeTypeAutoshapeExplosionOne = '\000j\000Y',
	WordMsoAutoShapeTypeAutoshapeExplosionTwo = '\000j\000Z',
	WordMsoAutoShapeTypeAutoshapeFourPointStar = '\000j\000[',
	WordMsoAutoShapeTypeAutoshapeFivePointStar = '\000j\000\\',
	WordMsoAutoShapeTypeAutoshapeEightPointStar = '\000j\000]',
	WordMsoAutoShapeTypeAutoshapeSixteenPointStar = '\000j\000^',
	WordMsoAutoShapeTypeAutoshapeTwentyFourPointStar = '\000j\000_',
	WordMsoAutoShapeTypeAutoshapeThirtyTwoPointStar = '\000j\000`',
	WordMsoAutoShapeTypeAutoshapeUpRibbon = '\000j\000a',
	WordMsoAutoShapeTypeAutoshapeDownRibbon = '\000j\000b',
	WordMsoAutoShapeTypeAutoshapeCurvedUpRibbon = '\000j\000c',
	WordMsoAutoShapeTypeAutoshapeCurvedDownRibbon = '\000j\000d',
	WordMsoAutoShapeTypeAutoshapeVerticalScroll = '\000j\000e',
	WordMsoAutoShapeTypeAutoshapeHorizontalScroll = '\000j\000f',
	WordMsoAutoShapeTypeAutoshapeWave = '\000j\000g',
	WordMsoAutoShapeTypeAutoshapeDoubleWave = '\000j\000h',
	WordMsoAutoShapeTypeAutoshapeRectangularCallout = '\000j\000i',
	WordMsoAutoShapeTypeAutoshapeRoundedRectangularCallout = '\000j\000j',
	WordMsoAutoShapeTypeAutoshapeOvalCallout = '\000j\000k',
	WordMsoAutoShapeTypeAutoshapeCloudCallout = '\000j\000l',
	WordMsoAutoShapeTypeAutoshapeLineCalloutOne = '\000j\000m',
	WordMsoAutoShapeTypeAutoshapeLineCalloutTwo = '\000j\000n',
	WordMsoAutoShapeTypeAutoshapeLineCalloutThree = '\000j\000o',
	WordMsoAutoShapeTypeAutoshapeLineCalloutFour = '\000j\000p',
	WordMsoAutoShapeTypeAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	WordMsoAutoShapeTypeAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	WordMsoAutoShapeTypeAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	WordMsoAutoShapeTypeAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	WordMsoAutoShapeTypeAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	WordMsoAutoShapeTypeAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	WordMsoAutoShapeTypeAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	WordMsoAutoShapeTypeAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	WordMsoAutoShapeTypeAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	WordMsoAutoShapeTypeAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	WordMsoAutoShapeTypeAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	WordMsoAutoShapeTypeAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	WordMsoAutoShapeTypeAutoshapeActionButtonCustom = '\000j\000}',
	WordMsoAutoShapeTypeAutoshapeActionButtonHome = '\000j\000~',
	WordMsoAutoShapeTypeAutoshapeActionButtonHelp = '\000j\000\177',
	WordMsoAutoShapeTypeAutoshapeActionButtonInformation = '\000j\000\200',
	WordMsoAutoShapeTypeAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	WordMsoAutoShapeTypeAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	WordMsoAutoShapeTypeAutoshapeActionButtonBeginning = '\000j\000\203',
	WordMsoAutoShapeTypeAutoshapeActionButtonEnd = '\000j\000\204',
	WordMsoAutoShapeTypeAutoshapeActionButtonReturn = '\000j\000\205',
	WordMsoAutoShapeTypeAutoshapeActionButtonDocument = '\000j\000\206',
	WordMsoAutoShapeTypeAutoshapeActionButtonSound = '\000j\000\207',
	WordMsoAutoShapeTypeAutoshapeActionButtonMovie = '\000j\000\210',
	WordMsoAutoShapeTypeAutoshapeBalloon = '\000j\000\211',
	WordMsoAutoShapeTypeAutoshapeNotPrimitive = '\000j\000\212',
	WordMsoAutoShapeTypeAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	WordMsoAutoShapeTypeAutoshapeLeftRightRibbon = '\000j\000\214',
	WordMsoAutoShapeTypeAutoshapeDiagonalStripe = '\000j\000\215',
	WordMsoAutoShapeTypeAutoshapePie = '\000j\000\216',
	WordMsoAutoShapeTypeAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	WordMsoAutoShapeTypeAutoshapeDecagon = '\000j\000\220',
	WordMsoAutoShapeTypeAutoshapeHeptagon = '\000j\000\221',
	WordMsoAutoShapeTypeAutoshapeDodecagon = '\000j\000\222',
	WordMsoAutoShapeTypeAutoshapeSixPointsStar = '\000j\000\223',
	WordMsoAutoShapeTypeAutoshapeSevenPointsStar = '\000j\000\224',
	WordMsoAutoShapeTypeAutoshapeTenPointsStar = '\000j\000\225',
	WordMsoAutoShapeTypeAutoshapeTwelvePointsStar = '\000j\000\226',
	WordMsoAutoShapeTypeAutoshapeRoundOneRectangle = '\000j\000\227',
	WordMsoAutoShapeTypeAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	WordMsoAutoShapeTypeAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	WordMsoAutoShapeTypeAutoshapeSnipRoundRectangle = '\000j\000\232',
	WordMsoAutoShapeTypeAutoshapeSnipOneRectangle = '\000j\000\233',
	WordMsoAutoShapeTypeAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	WordMsoAutoShapeTypeAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	WordMsoAutoShapeTypeAutoshapeFrame = '\000j\000\236',
	WordMsoAutoShapeTypeAutoshapeHalfFrame = '\000j\000\237',
	WordMsoAutoShapeTypeAutoshapeTear = '\000j\000\240',
	WordMsoAutoShapeTypeAutoshapeChord = '\000j\000\241',
	WordMsoAutoShapeTypeAutoshapeCorner = '\000j\000\242',
	WordMsoAutoShapeTypeAutoshapeMathPlus = '\000j\000\243',
	WordMsoAutoShapeTypeAutoshapeMathMinus = '\000j\000\244',
	WordMsoAutoShapeTypeAutoshapeMathMultiply = '\000j\000\245',
	WordMsoAutoShapeTypeAutoshapeMathDivide = '\000j\000\246',
	WordMsoAutoShapeTypeAutoshapeMathEqual = '\000j\000\247',
	WordMsoAutoShapeTypeAutoshapeMathNotEqual = '\000j\000\250',
	WordMsoAutoShapeTypeAutoshapeCornerTabs = '\000j\000\251',
	WordMsoAutoShapeTypeAutoshapeSquareTabs = '\000j\000\252',
	WordMsoAutoShapeTypeAutoshapePlaqueTabs = '\000j\000\253',
	WordMsoAutoShapeTypeAutoshapeGearSix = '\000j\000\254',
	WordMsoAutoShapeTypeAutoshapeGearNine = '\000j\000\255',
	WordMsoAutoShapeTypeAutoshapeFunnel = '\000j\000\256',
	WordMsoAutoShapeTypeAutoshapePieWedge = '\000j\000\257',
	WordMsoAutoShapeTypeAutoshapeLeftCircularArrow = '\000j\000\260',
	WordMsoAutoShapeTypeAutoshapeLeftRightCircularArrow = '\000j\000\261',
	WordMsoAutoShapeTypeAutoshapeSwooshArrow = '\000j\000\262',
	WordMsoAutoShapeTypeAutoshapeCloud = '\000j\000\263',
	WordMsoAutoShapeTypeAutoshapeChartX = '\000j\000\264',
	WordMsoAutoShapeTypeAutoshapeChartStar = '\000j\000\265',
	WordMsoAutoShapeTypeAutoshapeChartPlus = '\000j\000\266',
	WordMsoAutoShapeTypeAutoshapeLineInverse = '\000j\000\267'
};
typedef enum WordMsoAutoShapeType WordMsoAutoShapeType;

enum WordMsoShapeType {
	WordMsoShapeTypeShapeTypeUnset = '\000\213\377\376',
	WordMsoShapeTypeShapeTypeAuto = '\000\214\000\001',
	WordMsoShapeTypeShapeTypeCallout = '\000\214\000\002',
	WordMsoShapeTypeShapeTypeChart = '\000\214\000\003',
	WordMsoShapeTypeShapeTypeComment = '\000\214\000\004',
	WordMsoShapeTypeShapeTypeFreeForm = '\000\214\000\005',
	WordMsoShapeTypeShapeTypeGroup = '\000\214\000\006',
	WordMsoShapeTypeShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	WordMsoShapeTypeShapeTypeFormControl = '\000\214\000\010',
	WordMsoShapeTypeShapeTypeLine = '\000\214\000\011',
	WordMsoShapeTypeShapeTypeLinkedOLEObject = '\000\214\000\012',
	WordMsoShapeTypeShapeTypeLinkedPicture = '\000\214\000\013',
	WordMsoShapeTypeShapeTypeOLEControl = '\000\214\000\014',
	WordMsoShapeTypeShapeTypePicture = '\000\214\000\015',
	WordMsoShapeTypeShapeTypePlaceHolder = '\000\214\000\016',
	WordMsoShapeTypeShapeTypeWordArt = '\000\214\000\017',
	WordMsoShapeTypeShapeTypeMedia = '\000\214\000\020',
	WordMsoShapeTypeShapeTypeTextBox = '\000\214\000\021',
	WordMsoShapeTypeShapeTypeScriptAnchor = '\000\214\000\022',
	WordMsoShapeTypeShapeTypeTable = '\000\214\000\023',
	WordMsoShapeTypeShapeTypeCanvas = '\000\214\000\024',
	WordMsoShapeTypeShapeTypeDiagram = '\000\214\000\025',
	WordMsoShapeTypeShapeTypeInk = '\000\214\000\026',
	WordMsoShapeTypeShapeTypeInkComment = '\000\214\000\027',
	WordMsoShapeTypeShapeTypeSmartartGraphic = '\000\214\000\030',
	WordMsoShapeTypeShapeTypeSlicer = '\000\214\000\031',
	WordMsoShapeTypeShapeTypeWebVideo = '\000\214\000\032',
	WordMsoShapeTypeShapeTypeContentApplication = '\000\214\000\033',
	WordMsoShapeTypeShapeTypeGraphic = '\000\214\000\034',
	WordMsoShapeTypeShapeTypeLinkedGraphic = '\000\214\000\035',
	WordMsoShapeTypeShapeType3dModel = '\000\214\000\036',
	WordMsoShapeTypeShapeTypeLinked3dModel = '\000\214\000\037'
};
typedef enum WordMsoShapeType WordMsoShapeType;

enum WordMsoColorType {
	WordMsoColorTypeColorTypeUnset = '\000j\377\376',
	WordMsoColorTypeRGB = '\000k\000\001',
	WordMsoColorTypeScheme = '\000k\000\002'
};
typedef enum WordMsoColorType WordMsoColorType;

enum WordMsoPictureColorType {
	WordMsoPictureColorTypePictureColorTypeUnset = '\000\265\377\376',
	WordMsoPictureColorTypePictureColorAutomatic = '\000\266\000\001',
	WordMsoPictureColorTypePictureColorGrayScale = '\000\266\000\002',
	WordMsoPictureColorTypePictureColorBlackAndWhite = '\000\266\000\003',
	WordMsoPictureColorTypePictureColorWatermark = '\000\266\000\004'
};
typedef enum WordMsoPictureColorType WordMsoPictureColorType;

enum WordMsoCalloutAngleType {
	WordMsoCalloutAngleTypeAngleUnset = '\000k\377\376',
	WordMsoCalloutAngleTypeAngleAutomatic = '\000l\000\001',
	WordMsoCalloutAngleTypeAngle30 = '\000l\000\002',
	WordMsoCalloutAngleTypeAngle45 = '\000l\000\003',
	WordMsoCalloutAngleTypeAngle60 = '\000l\000\004',
	WordMsoCalloutAngleTypeAngle90 = '\000l\000\005'
};
typedef enum WordMsoCalloutAngleType WordMsoCalloutAngleType;

enum WordMsoCalloutDropType {
	WordMsoCalloutDropTypeDropUnset = '\000l\377\376',
	WordMsoCalloutDropTypeDropCustom = '\000m\000\001',
	WordMsoCalloutDropTypeDropTop = '\000m\000\002',
	WordMsoCalloutDropTypeDropCenter = '\000m\000\003',
	WordMsoCalloutDropTypeDropBottom = '\000m\000\004'
};
typedef enum WordMsoCalloutDropType WordMsoCalloutDropType;

enum WordMsoCalloutType {
	WordMsoCalloutTypeCalloutUnset = '\000m\377\376',
	WordMsoCalloutTypeCalloutOne = '\000n\000\001',
	WordMsoCalloutTypeCalloutTwo = '\000n\000\002',
	WordMsoCalloutTypeCalloutThree = '\000n\000\003',
	WordMsoCalloutTypeCalloutFour = '\000n\000\004'
};
typedef enum WordMsoCalloutType WordMsoCalloutType;

enum WordMsoTextOrientation {
	WordMsoTextOrientationTextOrientationUnset = '\000\215\377\376',
	WordMsoTextOrientationHorizontal = '\000\216\000\001',
	WordMsoTextOrientationUpward = '\000\216\000\002',
	WordMsoTextOrientationDownward = '\000\216\000\003',
	WordMsoTextOrientationVerticalEastAsian = '\000\216\000\004',
	WordMsoTextOrientationVertical = '\000\216\000\005',
	WordMsoTextOrientationHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum WordMsoTextOrientation WordMsoTextOrientation;

enum WordMsoScaleFrom {
	WordMsoScaleFromScaleFromTopLeft = '\000o\000\000',
	WordMsoScaleFromScaleFromMiddle = '\000o\000\001',
	WordMsoScaleFromScaleFromBottomRight = '\000o\000\002'
};
typedef enum WordMsoScaleFrom WordMsoScaleFrom;

enum WordMsoPresetCamera {
	WordMsoPresetCameraPresetCameraUnset = '\000\256\377\376',
	WordMsoPresetCameraCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	WordMsoPresetCameraCameraLegacyObliqueFromTop = '\000\257\000\002',
	WordMsoPresetCameraCameraLegacyObliqueFromTopright = '\000\257\000\003',
	WordMsoPresetCameraCameraLegacyObliqueFromLeft = '\000\257\000\004',
	WordMsoPresetCameraCameraLegacyObliqueFromFront = '\000\257\000\005',
	WordMsoPresetCameraCameraLegacyObliqueFromRight = '\000\257\000\006',
	WordMsoPresetCameraCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	WordMsoPresetCameraCameraLegacyObliqueFromBottom = '\000\257\000\010',
	WordMsoPresetCameraCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	WordMsoPresetCameraCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	WordMsoPresetCameraCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	WordMsoPresetCameraCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	WordMsoPresetCameraCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	WordMsoPresetCameraCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	WordMsoPresetCameraCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	WordMsoPresetCameraCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	WordMsoPresetCameraCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	WordMsoPresetCameraCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	WordMsoPresetCameraCameraOrthographic = '\000\257\000\023',
	WordMsoPresetCameraCameraIsometricFromTopUp = '\000\257\000\024',
	WordMsoPresetCameraCameraIsometricFromTopDown = '\000\257\000\025',
	WordMsoPresetCameraCameraIsometricFromBottomUp = '\000\257\000\026',
	WordMsoPresetCameraCameraIsometricFromBottomDown = '\000\257\000\027',
	WordMsoPresetCameraCameraIsometricFromLeftUp = '\000\257\000\030',
	WordMsoPresetCameraCameraIsometricFromLeftDown = '\000\257\000\031',
	WordMsoPresetCameraCameraIsometricFromRightUp = '\000\257\000\032',
	WordMsoPresetCameraCameraIsometricFromRightDown = '\000\257\000\033',
	WordMsoPresetCameraCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	WordMsoPresetCameraCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	WordMsoPresetCameraCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	WordMsoPresetCameraCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	WordMsoPresetCameraCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	WordMsoPresetCameraCameraIsometricOffAxis2FromTop = '\000\257\000!',
	WordMsoPresetCameraCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	WordMsoPresetCameraCameraIsometricOffAxis3FromRight = '\000\257\000#',
	WordMsoPresetCameraCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	WordMsoPresetCameraCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	WordMsoPresetCameraCameraIsometricOffAxis4FromRight = '\000\257\000&',
	WordMsoPresetCameraCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	WordMsoPresetCameraCameraObliqueFromTopLeft = '\000\257\000(',
	WordMsoPresetCameraCameraObliqueFromTop = '\000\257\000)',
	WordMsoPresetCameraCameraObliqueFromTopRight = '\000\257\000*',
	WordMsoPresetCameraCameraObliqueFromLeft = '\000\257\000+',
	WordMsoPresetCameraCameraObliqueFromRight = '\000\257\000,',
	WordMsoPresetCameraCameraObliqueFromBottomLeft = '\000\257\000-',
	WordMsoPresetCameraCameraObliqueFromBottom = '\000\257\000.',
	WordMsoPresetCameraCameraObliqueFromBottomRight = '\000\257\000/',
	WordMsoPresetCameraCameraPerspectiveFromFront = '\000\257\0000',
	WordMsoPresetCameraCameraPerspectiveFromLeft = '\000\257\0001',
	WordMsoPresetCameraCameraPerspectiveFromRight = '\000\257\0002',
	WordMsoPresetCameraCameraPerspectiveFromAbove = '\000\257\0003',
	WordMsoPresetCameraCameraPerspectiveFromBelow = '\000\257\0004',
	WordMsoPresetCameraCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	WordMsoPresetCameraCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	WordMsoPresetCameraCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	WordMsoPresetCameraCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	WordMsoPresetCameraCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	WordMsoPresetCameraCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	WordMsoPresetCameraCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	WordMsoPresetCameraCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	WordMsoPresetCameraCameraPerspectiveRelaxed = '\000\257\000=',
	WordMsoPresetCameraCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum WordMsoPresetCamera WordMsoPresetCamera;

enum WordMsoLightRigType {
	WordMsoLightRigTypeLightRigUnset = '\000\257\377\376',
	WordMsoLightRigTypeLightRigFlat1 = '\000\260\000\001',
	WordMsoLightRigTypeLightRigFlat2 = '\000\260\000\002',
	WordMsoLightRigTypeLightRigFlat3 = '\000\260\000\003',
	WordMsoLightRigTypeLightRigFlat4 = '\000\260\000\004',
	WordMsoLightRigTypeLightRigNormal1 = '\000\260\000\005',
	WordMsoLightRigTypeLightRigNormal2 = '\000\260\000\006',
	WordMsoLightRigTypeLightRigNormal3 = '\000\260\000\007',
	WordMsoLightRigTypeLightRigNormal4 = '\000\260\000\010',
	WordMsoLightRigTypeLightRigHarsh1 = '\000\260\000\011',
	WordMsoLightRigTypeLightRigHarsh2 = '\000\260\000\012',
	WordMsoLightRigTypeLightRigHarsh3 = '\000\260\000\013',
	WordMsoLightRigTypeLightRigHarsh4 = '\000\260\000\014',
	WordMsoLightRigTypeLightRigThreePoint = '\000\260\000\015',
	WordMsoLightRigTypeLightRigBalanced = '\000\260\000\016',
	WordMsoLightRigTypeLightRigSoft = '\000\260\000\017',
	WordMsoLightRigTypeLightRigHarsh = '\000\260\000\020',
	WordMsoLightRigTypeLightRigFlood = '\000\260\000\021',
	WordMsoLightRigTypeLightRigContrasting = '\000\260\000\022',
	WordMsoLightRigTypeLightRigMorning = '\000\260\000\023',
	WordMsoLightRigTypeLightRigSunrise = '\000\260\000\024',
	WordMsoLightRigTypeLightRigSunset = '\000\260\000\025',
	WordMsoLightRigTypeLightRigChilly = '\000\260\000\026',
	WordMsoLightRigTypeLightRigFreezing = '\000\260\000\027',
	WordMsoLightRigTypeLightRigFlat = '\000\260\000\030',
	WordMsoLightRigTypeLightRigTwoPoint = '\000\260\000\031',
	WordMsoLightRigTypeLightRigGlow = '\000\260\000\032',
	WordMsoLightRigTypeLightRigBrightRoom = '\000\260\000\033'
};
typedef enum WordMsoLightRigType WordMsoLightRigType;

enum WordMsoBevelType {
	WordMsoBevelTypeBevelTypeUnset = '\000\260\377\376',
	WordMsoBevelTypeBevelNone = '\000\261\000\001',
	WordMsoBevelTypeBevelRelaxedInset = '\000\261\000\002',
	WordMsoBevelTypeBevelCircle = '\000\261\000\003',
	WordMsoBevelTypeBevelSlope = '\000\261\000\004',
	WordMsoBevelTypeBevelCross = '\000\261\000\005',
	WordMsoBevelTypeBevelAngle = '\000\261\000\006',
	WordMsoBevelTypeBevelSoftRound = '\000\261\000\007',
	WordMsoBevelTypeBevelConvex = '\000\261\000\010',
	WordMsoBevelTypeBevelCoolSlant = '\000\261\000\011',
	WordMsoBevelTypeBevelDivot = '\000\261\000\012',
	WordMsoBevelTypeBevelRiblet = '\000\261\000\013',
	WordMsoBevelTypeBevelHardEdge = '\000\261\000\014',
	WordMsoBevelTypeBevelArtDeco = '\000\261\000\015'
};
typedef enum WordMsoBevelType WordMsoBevelType;

enum WordMsoShadowStyle {
	WordMsoShadowStyleShadowStyleUnset = '\000\261\377\376',
	WordMsoShadowStyleShadowStyleInner = '\000\262\000\001',
	WordMsoShadowStyleShadowStyleOuter = '\000\262\000\002'
};
typedef enum WordMsoShadowStyle WordMsoShadowStyle;

enum WordMsoParagraphAlignment {
	WordMsoParagraphAlignmentParagraphAlignmentUnset = '\000\346\377\376',
	WordMsoParagraphAlignmentParagraphAlignLeft = '\000\347\000\001',
	WordMsoParagraphAlignmentParagraphAlignCenter = '\000\347\000\002',
	WordMsoParagraphAlignmentParagraphAlignRight = '\000\347\000\003',
	WordMsoParagraphAlignmentParagraphAlignJustify = '\000\347\000\004',
	WordMsoParagraphAlignmentParagraphAlignDistribute = '\000\347\000\005',
	WordMsoParagraphAlignmentParagraphAlignThai = '\000\347\000\006',
	WordMsoParagraphAlignmentParagraphAlignJustifyLow = '\000\347\000\007'
};
typedef enum WordMsoParagraphAlignment WordMsoParagraphAlignment;

enum WordMsoTextStrike {
	WordMsoTextStrikeStrikeUnset = '\000\263\377\376',
	WordMsoTextStrikeNoStrike = '\000\264\000\000',
	WordMsoTextStrikeSingleStrike = '\000\264\000\001',
	WordMsoTextStrikeDoubleStrike = '\000\264\000\002'
};
typedef enum WordMsoTextStrike WordMsoTextStrike;

enum WordMsoTextCaps {
	WordMsoTextCapsCapsUnset = '\000\264\377\376',
	WordMsoTextCapsNoCaps = '\000\265\000\000',
	WordMsoTextCapsSmallCaps = '\000\265\000\001',
	WordMsoTextCapsAllCaps = '\000\265\000\002'
};
typedef enum WordMsoTextCaps WordMsoTextCaps;

enum WordMsoTextUnderlineType {
	WordMsoTextUnderlineTypeUnderlineUnset = '\003\356\377\376',
	WordMsoTextUnderlineTypeNoUnderline = '\003\357\000\000',
	WordMsoTextUnderlineTypeUnderlineWordsOnly = '\003\357\000\001',
	WordMsoTextUnderlineTypeUnderlineSingleLine = '\003\357\000\002',
	WordMsoTextUnderlineTypeUnderlineDoubleLine = '\003\357\000\003',
	WordMsoTextUnderlineTypeUnderlineHeavyLine = '\003\357\000\004',
	WordMsoTextUnderlineTypeUnderlineDottedLine = '\003\357\000\005',
	WordMsoTextUnderlineTypeUnderlineHeavyDottedLine = '\003\357\000\006',
	WordMsoTextUnderlineTypeUnderlineDashLine = '\003\357\000\007',
	WordMsoTextUnderlineTypeUnderlineHeavyDashLine = '\003\357\000\010',
	WordMsoTextUnderlineTypeUnderlineLongDashLine = '\003\357\000\011',
	WordMsoTextUnderlineTypeUnderlineHeavyLongDashLine = '\003\357\000\012',
	WordMsoTextUnderlineTypeUnderlineDotDashLine = '\003\357\000\013',
	WordMsoTextUnderlineTypeUnderlineHeavyDotDashLine = '\003\357\000\014',
	WordMsoTextUnderlineTypeUnderlineDotDotDashLine = '\003\357\000\015',
	WordMsoTextUnderlineTypeUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	WordMsoTextUnderlineTypeUnderlineWavyLine = '\003\357\000\017',
	WordMsoTextUnderlineTypeUnderlineHeavyWavyLine = '\003\357\000\020',
	WordMsoTextUnderlineTypeUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum WordMsoTextUnderlineType WordMsoTextUnderlineType;

enum WordMsoTextTabAlign {
	WordMsoTextTabAlignTabUnset = '\000\266\377\376',
	WordMsoTextTabAlignLeftTab = '\000\267\000\000',
	WordMsoTextTabAlignCenterTab = '\000\267\000\001',
	WordMsoTextTabAlignRightTab = '\000\267\000\002',
	WordMsoTextTabAlignDecimalTab = '\000\267\000\003'
};
typedef enum WordMsoTextTabAlign WordMsoTextTabAlign;

enum WordMsoTextCharWrap {
	WordMsoTextCharWrapCharacterWrapUnset = '\000\267\377\376',
	WordMsoTextCharWrapNoCharacterWrap = '\000\270\000\000',
	WordMsoTextCharWrapStandardCharacterWrap = '\000\270\000\001',
	WordMsoTextCharWrapStrictCharacterWrap = '\000\270\000\002',
	WordMsoTextCharWrapCustomCharacterWrap = '\000\270\000\003'
};
typedef enum WordMsoTextCharWrap WordMsoTextCharWrap;

enum WordMsoTextFontAlign {
	WordMsoTextFontAlignFontAlignmentUnset = '\000\270\377\376',
	WordMsoTextFontAlignAutomaticAlignment = '\000\271\000\000',
	WordMsoTextFontAlignTopAlignment = '\000\271\000\001',
	WordMsoTextFontAlignCenterAlignment = '\000\271\000\002',
	WordMsoTextFontAlignBaselineAlignment = '\000\271\000\003',
	WordMsoTextFontAlignBottomAlignment = '\000\271\000\004'
};
typedef enum WordMsoTextFontAlign WordMsoTextFontAlign;

enum WordMsoAutoSize {
	WordMsoAutoSizeAutoSizeUnset = '\000\344\377\376',
	WordMsoAutoSizeAutoSizeNone = '\000\345\000\000',
	WordMsoAutoSizeShapeToFitText = '\000\345\000\001',
	WordMsoAutoSizeTextToFitShape = '\000\345\000\002'
};
typedef enum WordMsoAutoSize WordMsoAutoSize;

enum WordMsoPathFormat {
	WordMsoPathFormatPathTypeUnset = '\000\272\377\376',
	WordMsoPathFormatNoPathType = '\000\273\000\000',
	WordMsoPathFormatPathType1 = '\000\273\000\001',
	WordMsoPathFormatPathType2 = '\000\273\000\002',
	WordMsoPathFormatPathType3 = '\000\273\000\003',
	WordMsoPathFormatPathType4 = '\000\273\000\004'
};
typedef enum WordMsoPathFormat WordMsoPathFormat;

enum WordMsoWarpFormat {
	WordMsoWarpFormatWarpFormatUnset = '\000\273\377\376',
	WordMsoWarpFormatWarpFormat1 = '\000\274\000\000',
	WordMsoWarpFormatWarpFormat2 = '\000\274\000\001',
	WordMsoWarpFormatWarpFormat3 = '\000\274\000\002',
	WordMsoWarpFormatWarpFormat4 = '\000\274\000\003',
	WordMsoWarpFormatWarpFormat5 = '\000\274\000\004',
	WordMsoWarpFormatWarpFormat6 = '\000\274\000\005',
	WordMsoWarpFormatWarpFormat7 = '\000\274\000\006',
	WordMsoWarpFormatWarpFormat8 = '\000\274\000\007',
	WordMsoWarpFormatWarpFormat9 = '\000\274\000\010',
	WordMsoWarpFormatWarpFormat10 = '\000\274\000\011',
	WordMsoWarpFormatWarpFormat11 = '\000\274\000\012',
	WordMsoWarpFormatWarpFormat12 = '\000\274\000\013',
	WordMsoWarpFormatWarpFormat13 = '\000\274\000\014',
	WordMsoWarpFormatWarpFormat14 = '\000\274\000\015',
	WordMsoWarpFormatWarpFormat15 = '\000\274\000\016',
	WordMsoWarpFormatWarpFormat16 = '\000\274\000\017',
	WordMsoWarpFormatWarpFormat17 = '\000\274\000\020',
	WordMsoWarpFormatWarpFormat18 = '\000\274\000\021',
	WordMsoWarpFormatWarpFormat19 = '\000\274\000\022',
	WordMsoWarpFormatWarpFormat20 = '\000\274\000\023',
	WordMsoWarpFormatWarpFormat21 = '\000\274\000\024',
	WordMsoWarpFormatWarpFormat22 = '\000\274\000\025',
	WordMsoWarpFormatWarpFormat23 = '\000\274\000\026',
	WordMsoWarpFormatWarpFormat24 = '\000\274\000\027',
	WordMsoWarpFormatWarpFormat25 = '\000\274\000\030',
	WordMsoWarpFormatWarpFormat26 = '\000\274\000\031',
	WordMsoWarpFormatWarpFormat27 = '\000\274\000\032',
	WordMsoWarpFormatWarpFormat28 = '\000\274\000\033',
	WordMsoWarpFormatWarpFormat29 = '\000\274\000\034',
	WordMsoWarpFormatWarpFormat30 = '\000\274\000\035',
	WordMsoWarpFormatWarpFormat31 = '\000\274\000\036',
	WordMsoWarpFormatWarpFormat32 = '\000\274\000\037',
	WordMsoWarpFormatWarpFormat33 = '\000\274\000 ',
	WordMsoWarpFormatWarpFormat34 = '\000\274\000!',
	WordMsoWarpFormatWarpFormat35 = '\000\274\000\"',
	WordMsoWarpFormatWarpFormat36 = '\000\274\000#',
	WordMsoWarpFormatWarpFormat37 = '\000\274\000$'
};
typedef enum WordMsoWarpFormat WordMsoWarpFormat;

enum WordMsoTextChangeCase {
	WordMsoTextChangeCaseCaseSentence = '\000\344\000\001',
	WordMsoTextChangeCaseCaseLower = '\000\344\000\002',
	WordMsoTextChangeCaseCaseUpper = '\000\344\000\003',
	WordMsoTextChangeCaseCaseTitle = '\000\344\000\004',
	WordMsoTextChangeCaseCaseToggle = '\000\344\000\005'
};
typedef enum WordMsoTextChangeCase WordMsoTextChangeCase;

enum WordMsoDateTimeFormat {
	WordMsoDateTimeFormatDateTimeFormatUnset = '\000\275\377\376',
	WordMsoDateTimeFormatDateTimeFormatMdyy = '\000\276\000\001',
	WordMsoDateTimeFormatDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	WordMsoDateTimeFormatDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	WordMsoDateTimeFormatDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	WordMsoDateTimeFormatDateTimeFormatDMMMyy = '\000\276\000\005',
	WordMsoDateTimeFormatDateTimeFormatMMMMyy = '\000\276\000\006',
	WordMsoDateTimeFormatDateTimeFormatMMyy = '\000\276\000\007',
	WordMsoDateTimeFormatDateTimeFormatMMddyyHmm = '\000\276\000\010',
	WordMsoDateTimeFormatDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	WordMsoDateTimeFormatDateTimeFormatHmm = '\000\276\000\012',
	WordMsoDateTimeFormatDateTimeFormatHmmss = '\000\276\000\013',
	WordMsoDateTimeFormatDateTimeFormatHmmAMPM = '\000\276\000\014',
	WordMsoDateTimeFormatDateTimeFormatHmmssAMPM = '\000\276\000\015',
	WordMsoDateTimeFormatDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum WordMsoDateTimeFormat WordMsoDateTimeFormat;

enum WordMsoSoftEdgeType {
	WordMsoSoftEdgeTypeSoftEdgeUnset = '\000\277\377\376',
	WordMsoSoftEdgeTypeNoSoftEdge = '\000\300\000\000',
	WordMsoSoftEdgeTypeSoftEdgeType1 = '\000\300\000\001',
	WordMsoSoftEdgeTypeSoftEdgeType2 = '\000\300\000\002',
	WordMsoSoftEdgeTypeSoftEdgeType3 = '\000\300\000\003',
	WordMsoSoftEdgeTypeSoftEdgeType4 = '\000\300\000\004',
	WordMsoSoftEdgeTypeSoftEdgeType5 = '\000\300\000\005',
	WordMsoSoftEdgeTypeSoftEdgeType6 = '\000\300\000\006'
};
typedef enum WordMsoSoftEdgeType WordMsoSoftEdgeType;

enum WordMsoThemeColorSchemeIndex {
	WordMsoThemeColorSchemeIndexFirstDarkSchemeColor = '\000\301\000\001',
	WordMsoThemeColorSchemeIndexFirstLightSchemeColor = '\000\301\000\002',
	WordMsoThemeColorSchemeIndexSecondDarkSchemeColor = '\000\301\000\003',
	WordMsoThemeColorSchemeIndexSecondLightSchemeColor = '\000\301\000\004',
	WordMsoThemeColorSchemeIndexFirstAccentSchemeColor = '\000\301\000\005',
	WordMsoThemeColorSchemeIndexSecondAccentSchemeColor = '\000\301\000\006',
	WordMsoThemeColorSchemeIndexThirdAccentSchemeColor = '\000\301\000\007',
	WordMsoThemeColorSchemeIndexFourthAccentSchemeColor = '\000\301\000\010',
	WordMsoThemeColorSchemeIndexFifthAccentSchemeColor = '\000\301\000\011',
	WordMsoThemeColorSchemeIndexSixthAccentSchemeColor = '\000\301\000\012',
	WordMsoThemeColorSchemeIndexHyperlinkSchemeColor = '\000\301\000\013',
	WordMsoThemeColorSchemeIndexFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum WordMsoThemeColorSchemeIndex WordMsoThemeColorSchemeIndex;

enum WordMsoThemeColorIndex {
	WordMsoThemeColorIndexThemeColorUnset = '\000\301\377\376',
	WordMsoThemeColorIndexNoThemeColor = '\000\302\000\000',
	WordMsoThemeColorIndexFirstDarkThemeColor = '\000\302\000\001',
	WordMsoThemeColorIndexFirstLightThemeColor = '\000\302\000\002',
	WordMsoThemeColorIndexSecondDarkThemeColor = '\000\302\000\003',
	WordMsoThemeColorIndexSecondLightThemeColor = '\000\302\000\004',
	WordMsoThemeColorIndexFirstAccentThemeColor = '\000\302\000\005',
	WordMsoThemeColorIndexSecondAccentThemeColor = '\000\302\000\006',
	WordMsoThemeColorIndexThirdAccentThemeColor = '\000\302\000\007',
	WordMsoThemeColorIndexFourthAccentThemeColor = '\000\302\000\010',
	WordMsoThemeColorIndexFifthAccentThemeColor = '\000\302\000\011',
	WordMsoThemeColorIndexSixthAccentThemeColor = '\000\302\000\012',
	WordMsoThemeColorIndexHyperlinkThemeColor = '\000\302\000\013',
	WordMsoThemeColorIndexFollowedHyperlinkThemeColor = '\000\302\000\014',
	WordMsoThemeColorIndexFirstTextThemeColor = '\000\302\000\015',
	WordMsoThemeColorIndexFirstBackgroundThemeColor = '\000\302\000\016',
	WordMsoThemeColorIndexSecondTextThemeColor = '\000\302\000\017',
	WordMsoThemeColorIndexSecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum WordMsoThemeColorIndex WordMsoThemeColorIndex;

enum WordMsoFontLanguageIndex {
	WordMsoFontLanguageIndexThemeFontLatin = '\000\303\000\001',
	WordMsoFontLanguageIndexThemeFontComplexScript = '\000\303\000\002',
	WordMsoFontLanguageIndexThemeFontHighAnsi = '\000\303\000\003',
	WordMsoFontLanguageIndexThemeFontEastAsian = '\000\303\000\004'
};
typedef enum WordMsoFontLanguageIndex WordMsoFontLanguageIndex;

enum WordMsoShapeStyleIndex {
	WordMsoShapeStyleIndexShapeStyleUnset = '\000\303\377\376',
	WordMsoShapeStyleIndexShapeNotAPreset = '\000\304\000\000',
	WordMsoShapeStyleIndexShapePreset1 = '\000\304\000\001',
	WordMsoShapeStyleIndexShapePreset2 = '\000\304\000\002',
	WordMsoShapeStyleIndexShapePreset3 = '\000\304\000\003',
	WordMsoShapeStyleIndexShapePreset4 = '\000\304\000\004',
	WordMsoShapeStyleIndexShapePreset5 = '\000\304\000\005',
	WordMsoShapeStyleIndexShapePreset6 = '\000\304\000\006',
	WordMsoShapeStyleIndexShapePreset7 = '\000\304\000\007',
	WordMsoShapeStyleIndexShapePreset8 = '\000\304\000\010',
	WordMsoShapeStyleIndexShapePreset9 = '\000\304\000\011',
	WordMsoShapeStyleIndexShapePreset10 = '\000\304\000\012',
	WordMsoShapeStyleIndexShapePreset11 = '\000\304\000\013',
	WordMsoShapeStyleIndexShapePreset12 = '\000\304\000\014',
	WordMsoShapeStyleIndexShapePreset13 = '\000\304\000\015',
	WordMsoShapeStyleIndexShapePreset14 = '\000\304\000\016',
	WordMsoShapeStyleIndexShapePreset15 = '\000\304\000\017',
	WordMsoShapeStyleIndexShapePreset16 = '\000\304\000\020',
	WordMsoShapeStyleIndexShapePreset17 = '\000\304\000\021',
	WordMsoShapeStyleIndexShapePreset18 = '\000\304\000\022',
	WordMsoShapeStyleIndexShapePreset19 = '\000\304\000\023',
	WordMsoShapeStyleIndexShapePreset20 = '\000\304\000\024',
	WordMsoShapeStyleIndexShapePreset21 = '\000\304\000\025',
	WordMsoShapeStyleIndexShapePreset22 = '\000\304\000\026',
	WordMsoShapeStyleIndexShapePreset23 = '\000\304\000\027',
	WordMsoShapeStyleIndexShapePreset24 = '\000\304\000\030',
	WordMsoShapeStyleIndexShapePreset25 = '\000\304\000\031',
	WordMsoShapeStyleIndexShapePreset26 = '\000\304\000\032',
	WordMsoShapeStyleIndexShapePreset27 = '\000\304\000\033',
	WordMsoShapeStyleIndexShapePreset28 = '\000\304\000\034',
	WordMsoShapeStyleIndexShapePreset29 = '\000\304\000\035',
	WordMsoShapeStyleIndexShapePreset30 = '\000\304\000\036',
	WordMsoShapeStyleIndexShapePreset31 = '\000\304\000\037',
	WordMsoShapeStyleIndexShapePreset32 = '\000\304\000 ',
	WordMsoShapeStyleIndexShapePreset33 = '\000\304\000!',
	WordMsoShapeStyleIndexShapePreset34 = '\000\304\000\"',
	WordMsoShapeStyleIndexShapePreset35 = '\000\304\000#',
	WordMsoShapeStyleIndexShapePreset36 = '\000\304\000$',
	WordMsoShapeStyleIndexShapePreset37 = '\000\304\000%',
	WordMsoShapeStyleIndexShapePreset38 = '\000\304\000&',
	WordMsoShapeStyleIndexShapePreset39 = '\000\304\000\'',
	WordMsoShapeStyleIndexShapePreset40 = '\000\304\000(',
	WordMsoShapeStyleIndexShapePreset41 = '\000\304\000)',
	WordMsoShapeStyleIndexShapePreset42 = '\000\304\000*',
	WordMsoShapeStyleIndexLinePreset1 = '\000\304\'\021',
	WordMsoShapeStyleIndexLinePreset2 = '\000\304\'\022',
	WordMsoShapeStyleIndexLinePreset3 = '\000\304\'\023',
	WordMsoShapeStyleIndexLinePreset4 = '\000\304\'\024',
	WordMsoShapeStyleIndexLinePreset5 = '\000\304\'\025',
	WordMsoShapeStyleIndexLinePreset6 = '\000\304\'\026',
	WordMsoShapeStyleIndexLinePreset7 = '\000\304\'\027',
	WordMsoShapeStyleIndexLinePreset8 = '\000\304\'\030',
	WordMsoShapeStyleIndexLinePreset9 = '\000\304\'\031',
	WordMsoShapeStyleIndexLinePreset10 = '\000\304\'\032',
	WordMsoShapeStyleIndexLinePreset11 = '\000\304\'\033',
	WordMsoShapeStyleIndexLinePreset12 = '\000\304\'\034',
	WordMsoShapeStyleIndexLinePreset13 = '\000\304\'\035',
	WordMsoShapeStyleIndexLinePreset14 = '\000\304\'\036',
	WordMsoShapeStyleIndexLinePreset15 = '\000\304\'\037',
	WordMsoShapeStyleIndexLinePreset16 = '\000\304\' ',
	WordMsoShapeStyleIndexLinePreset17 = '\000\304\'!',
	WordMsoShapeStyleIndexLinePreset18 = '\000\304\'\"',
	WordMsoShapeStyleIndexLinePreset19 = '\000\304\'#',
	WordMsoShapeStyleIndexLinePreset20 = '\000\304\'$',
	WordMsoShapeStyleIndexLinePreset21 = '\000\304\'%'
};
typedef enum WordMsoShapeStyleIndex WordMsoShapeStyleIndex;

enum WordMsoBackgroundStyleIndex {
	WordMsoBackgroundStyleIndexBackgroundUnset = '\000\304\377\376',
	WordMsoBackgroundStyleIndexBackgroundNotAPreset = '\000\305\000\000',
	WordMsoBackgroundStyleIndexBackgroundPreset1 = '\000\305\000\001',
	WordMsoBackgroundStyleIndexBackgroundPreset2 = '\000\305\000\002',
	WordMsoBackgroundStyleIndexBackgroundPreset3 = '\000\305\000\003',
	WordMsoBackgroundStyleIndexBackgroundPreset4 = '\000\305\000\004',
	WordMsoBackgroundStyleIndexBackgroundPreset5 = '\000\305\000\005',
	WordMsoBackgroundStyleIndexBackgroundPreset6 = '\000\305\000\006',
	WordMsoBackgroundStyleIndexBackgroundPreset7 = '\000\305\000\007',
	WordMsoBackgroundStyleIndexBackgroundPreset8 = '\000\305\000\010',
	WordMsoBackgroundStyleIndexBackgroundPreset9 = '\000\305\000\011',
	WordMsoBackgroundStyleIndexBackgroundPreset10 = '\000\305\000\012',
	WordMsoBackgroundStyleIndexBackgroundPreset11 = '\000\305\000\013',
	WordMsoBackgroundStyleIndexBackgroundPreset12 = '\000\305\000\014'
};
typedef enum WordMsoBackgroundStyleIndex WordMsoBackgroundStyleIndex;

enum WordMsoTextDirection {
	WordMsoTextDirectionTextDirectionUnset = '\000\352\377\376',
	WordMsoTextDirectionLeftToRight = '\000\353\000\001',
	WordMsoTextDirectionRightToLeft = '\000\353\000\002'
};
typedef enum WordMsoTextDirection WordMsoTextDirection;

enum WordMsoBulletType {
	WordMsoBulletTypeBulletTypeUnset = '\000\347\377\376',
	WordMsoBulletTypeBulletTypeNone = '\000\350\000\000',
	WordMsoBulletTypeBulletTypeUnnumbered = '\000\350\000\001',
	WordMsoBulletTypeBulletTypeNumbered = '\000\350\000\002',
	WordMsoBulletTypePictureBulletType = '\000\350\000\003'
};
typedef enum WordMsoBulletType WordMsoBulletType;

enum WordMsoNumberedBulletStyle {
	WordMsoNumberedBulletStyleNumberedBulletStyleUnset = '\000\350\377\376',
	WordMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	WordMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	WordMsoNumberedBulletStyleNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	WordMsoNumberedBulletStyleNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	WordMsoNumberedBulletStyleNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	WordMsoNumberedBulletStyleNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	WordMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	WordMsoNumberedBulletStyleNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	WordMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	WordMsoNumberedBulletStyleNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicPlain = '\000\351\000\015',
	WordMsoNumberedBulletStyleNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	WordMsoNumberedBulletStyleNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	WordMsoNumberedBulletStyleNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	WordMsoNumberedBulletStyleNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	WordMsoNumberedBulletStyleNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	WordMsoNumberedBulletStyleNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	WordMsoNumberedBulletStyleNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	WordMsoNumberedBulletStyleNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	WordMsoNumberedBulletStyleNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	WordMsoNumberedBulletStyleNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	WordMsoNumberedBulletStyleNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	WordMsoNumberedBulletStyleNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	WordMsoNumberedBulletStyleNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	WordMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	WordMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	WordMsoNumberedBulletStyleNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	WordMsoNumberedBulletStyleNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	WordMsoNumberedBulletStyleNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	WordMsoNumberedBulletStyleNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	WordMsoNumberedBulletStyleNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	WordMsoNumberedBulletStyleNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	WordMsoNumberedBulletStyleNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	WordMsoNumberedBulletStyleNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	WordMsoNumberedBulletStyleNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum WordMsoNumberedBulletStyle WordMsoNumberedBulletStyle;

enum WordMsoTabStopType {
	WordMsoTabStopTypeTabstopUnset = '\000\364\377\376',
	WordMsoTabStopTypeTabstopLeft = '\000\365\000\001',
	WordMsoTabStopTypeTabstopCenter = '\000\365\000\002',
	WordMsoTabStopTypeTabstopRight = '\000\365\000\003',
	WordMsoTabStopTypeTabstopDecimal = '\000\365\000\004'
};
typedef enum WordMsoTabStopType WordMsoTabStopType;

enum WordMsoReflectionType {
	WordMsoReflectionTypeReflectionUnset = '\003\351\377\376',
	WordMsoReflectionTypeReflectionTypeNone = '\003\352\000\000',
	WordMsoReflectionTypeReflectionType1 = '\003\352\000\001',
	WordMsoReflectionTypeReflectionType2 = '\003\352\000\002',
	WordMsoReflectionTypeReflectionType3 = '\003\352\000\003',
	WordMsoReflectionTypeReflectionType4 = '\003\352\000\004',
	WordMsoReflectionTypeReflectionType5 = '\003\352\000\005',
	WordMsoReflectionTypeReflectionType6 = '\003\352\000\006',
	WordMsoReflectionTypeReflectionType7 = '\003\352\000\007',
	WordMsoReflectionTypeReflectionType8 = '\003\352\000\010',
	WordMsoReflectionTypeReflectionType9 = '\003\352\000\011'
};
typedef enum WordMsoReflectionType WordMsoReflectionType;

enum WordMsoTextureAlignment {
	WordMsoTextureAlignmentTextureUnset = '\003\352\377\376',
	WordMsoTextureAlignmentTextureTopLeft = '\003\353\000\000',
	WordMsoTextureAlignmentTextureTop = '\003\353\000\001',
	WordMsoTextureAlignmentTextureTopRight = '\003\353\000\002',
	WordMsoTextureAlignmentTextureLeft = '\003\353\000\003',
	WordMsoTextureAlignmentTextureCenter = '\003\353\000\004',
	WordMsoTextureAlignmentTextureRight = '\003\353\000\005',
	WordMsoTextureAlignmentTextureBottomLeft = '\003\353\000\006',
	WordMsoTextureAlignmentTextureBotton = '\003\353\000\007',
	WordMsoTextureAlignmentTextureBottomRight = '\003\353\000\010'
};
typedef enum WordMsoTextureAlignment WordMsoTextureAlignment;

enum WordMsoBaselineAlignment {
	WordMsoBaselineAlignmentTextBaselineAlignmentUnset = '\003\353\377\376',
	WordMsoBaselineAlignmentTextBaselineAlignBaseline = '\003\354\000\001',
	WordMsoBaselineAlignmentTextBaselineAlignTop = '\003\354\000\002',
	WordMsoBaselineAlignmentTextBaselineAlignCenter = '\003\354\000\003',
	WordMsoBaselineAlignmentTextBaselineAlignEastAsian50 = '\003\354\000\004',
	WordMsoBaselineAlignmentTextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum WordMsoBaselineAlignment WordMsoBaselineAlignment;

enum WordMsoClipboardFormat {
	WordMsoClipboardFormatClipboardFormatUnset = '\003\354\377\376',
	WordMsoClipboardFormatNativeClipboardFormat = '\003\355\000\001',
	WordMsoClipboardFormatHTMlClipboardFormat = '\003\355\000\002',
	WordMsoClipboardFormatRTFClipboardFormat = '\003\355\000\003',
	WordMsoClipboardFormatPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum WordMsoClipboardFormat WordMsoClipboardFormat;

enum WordMsoTextRangeInsertPosition {
	WordMsoTextRangeInsertPositionInsertBefore = '\003\356\000\000',
	WordMsoTextRangeInsertPositionInsertAfter = '\003\356\000\001'
};
typedef enum WordMsoTextRangeInsertPosition WordMsoTextRangeInsertPosition;

enum WordMsoPictureType {
	WordMsoPictureTypeSaveAsDefault = '\003\362\377\376',
	WordMsoPictureTypeSaveAsPNGFile = '\003\363\000\000',
	WordMsoPictureTypeSaveAsBMPFile = '\003\363\000\001',
	WordMsoPictureTypeSaveAsGIFFile = '\003\363\000\002',
	WordMsoPictureTypeSaveAsJPGFile = '\003\363\000\003',
	WordMsoPictureTypeSaveAsPDFFile = '\003\363\000\004'
};
typedef enum WordMsoPictureType WordMsoPictureType;

enum WordMsoPictureEffectType {
	WordMsoPictureEffectTypeNoEffect = '\003\364\000\000',
	WordMsoPictureEffectTypeEffectBackgroundRemoval = '\003\364\000\001',
	WordMsoPictureEffectTypeEffectBlur = '\003\364\000\002',
	WordMsoPictureEffectTypeEffectBrightnessContrast = '\003\364\000\003',
	WordMsoPictureEffectTypeEffectCement = '\003\364\000\004',
	WordMsoPictureEffectTypeEffectCrisscrossEtching = '\003\364\000\005',
	WordMsoPictureEffectTypeEffectChalkSketch = '\003\364\000\006',
	WordMsoPictureEffectTypeEffectColorTemperature = '\003\364\000\007',
	WordMsoPictureEffectTypeEffectCutout = '\003\364\000\010',
	WordMsoPictureEffectTypeEffectFilmGrain = '\003\364\000\011',
	WordMsoPictureEffectTypeEffectGlass = '\003\364\000\012',
	WordMsoPictureEffectTypeEffectGlowDiffused = '\003\364\000\013',
	WordMsoPictureEffectTypeEffectGlowEdges = '\003\364\000\014',
	WordMsoPictureEffectTypeEffectLightScreen = '\003\364\000\015',
	WordMsoPictureEffectTypeEffectLineDrawing = '\003\364\000\016',
	WordMsoPictureEffectTypeEffectMarker = '\003\364\000\017',
	WordMsoPictureEffectTypeEffectMosiaicBubbles = '\003\364\000\020',
	WordMsoPictureEffectTypeEffectPaintBrush = '\003\364\000\021',
	WordMsoPictureEffectTypeEffectPaintStrokes = '\003\364\000\022',
	WordMsoPictureEffectTypeEffectPastelsSmooth = '\003\364\000\023',
	WordMsoPictureEffectTypeEffectPencilGrayscale = '\003\364\000\024',
	WordMsoPictureEffectTypeEffectPencilSketch = '\003\364\000\025',
	WordMsoPictureEffectTypeEffectPhotocopy = '\003\364\000\026',
	WordMsoPictureEffectTypeEffectPlasticWrap = '\003\364\000\027',
	WordMsoPictureEffectTypeEffectSaturation = '\003\364\000\030',
	WordMsoPictureEffectTypeEffectSharpenSoften = '\003\364\000\031',
	WordMsoPictureEffectTypeEffectTexturizer = '\003\364\000\032',
	WordMsoPictureEffectTypeEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum WordMsoPictureEffectType WordMsoPictureEffectType;

enum WordMsoSegmentType {
	WordMsoSegmentTypeLine = '\000\217\000\000',
	WordMsoSegmentTypeCurve = '\000\217\000\001'
};
typedef enum WordMsoSegmentType WordMsoSegmentType;

enum WordMsoEditingType {
	WordMsoEditingTypeAuto = '\000\220\000\000',
	WordMsoEditingTypeCorner = '\000\220\000\001',
	WordMsoEditingTypeSmooth = '\000\220\000\002',
	WordMsoEditingTypeSymmetric = '\000\220\000\003'
};
typedef enum WordMsoEditingType WordMsoEditingType;

enum WordMsoSmartArtNodePosition {
	WordMsoSmartArtNodePositionDefaultNodePosition = '\003\365\000\001',
	WordMsoSmartArtNodePositionAfterNode = '\003\365\000\002',
	WordMsoSmartArtNodePositionBeforeNode = '\003\365\000\003',
	WordMsoSmartArtNodePositionAboveNode = '\003\365\000\004',
	WordMsoSmartArtNodePositionBelowNode = '\003\365\000\005'
};
typedef enum WordMsoSmartArtNodePosition WordMsoSmartArtNodePosition;

enum WordMsoSmartArtNodeType {
	WordMsoSmartArtNodeTypeDefaultNode = '\003\366\000\001',
	WordMsoSmartArtNodeTypeAssistantNode = '\003\366\000\002'
};
typedef enum WordMsoSmartArtNodeType WordMsoSmartArtNodeType;

enum WordMsoOrgChartLayoutType {
	WordMsoOrgChartLayoutTypeOrgChartLayoutUnset = '\003\366\377\376',
	WordMsoOrgChartLayoutTypeOrgChartLayoutStandard = '\003\367\000\001',
	WordMsoOrgChartLayoutTypeOrgChartLayoutBothHanging = '\003\367\000\002',
	WordMsoOrgChartLayoutTypeOrgChartLayoutLeftHanging = '\003\367\000\003',
	WordMsoOrgChartLayoutTypeOrgChartLayoutRightHanging = '\003\367\000\004',
	WordMsoOrgChartLayoutTypeOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum WordMsoOrgChartLayoutType WordMsoOrgChartLayoutType;

enum WordMsoAlignCmd {
	WordMsoAlignCmdAlignLefts = '\000\000\000\000',
	WordMsoAlignCmdAlignCenters = '\000\000\000\001',
	WordMsoAlignCmdAlignRights = '\000\000\000\002',
	WordMsoAlignCmdAlignTops = '\000\000\000\003',
	WordMsoAlignCmdAlignMiddles = '\000\000\000\004',
	WordMsoAlignCmdAlignBottoms = '\000\000\000\005'
};
typedef enum WordMsoAlignCmd WordMsoAlignCmd;

enum WordMsoDistributeCmd {
	WordMsoDistributeCmdDistributeHorizontally = '\000\000\000\000',
	WordMsoDistributeCmdDistributeVertically = '\000\000\000\001'
};
typedef enum WordMsoDistributeCmd WordMsoDistributeCmd;

enum WordMsoOrientation {
	WordMsoOrientationOrientationUnset = '\000\214\377\376',
	WordMsoOrientationHorizontalOrientation = '\000\215\000\001',
	WordMsoOrientationVerticalOrientation = '\000\215\000\002'
};
typedef enum WordMsoOrientation WordMsoOrientation;

enum WordMsoZOrderCmd {
	WordMsoZOrderCmdBringShapeToFront = '\000p\000\000',
	WordMsoZOrderCmdSendShapeToBack = '\000p\000\001',
	WordMsoZOrderCmdBringShapeForward = '\000p\000\002',
	WordMsoZOrderCmdSendShapeBackward = '\000p\000\003',
	WordMsoZOrderCmdBringShapeInFrontOfText = '\000p\000\004',
	WordMsoZOrderCmdSendShapeBehindText = '\000p\000\005'
};
typedef enum WordMsoZOrderCmd WordMsoZOrderCmd;

enum WordMsoScriptLanguage {
	WordMsoScriptLanguageBringShapeToFront = '\000o\000\001',
	WordMsoScriptLanguageSendShapeToBack = '\000o\000\002',
	WordMsoScriptLanguageBringShapeForward = '\000o\000\003',
	WordMsoScriptLanguageSendShapeBackward = '\000o\000\004'
};
typedef enum WordMsoScriptLanguage WordMsoScriptLanguage;

enum WordMsoFlipCmd {
	WordMsoFlipCmdFlipHorizontal = '\000q\000\000',
	WordMsoFlipCmdFlipVertical = '\000q\000\001'
};
typedef enum WordMsoFlipCmd WordMsoFlipCmd;

enum WordMsoTriState {
	WordMsoTriStateTrue = '\000\240\377\377',
	WordMsoTriStateFalse = '\000\241\000\000',
	WordMsoTriStateCTrue = '\000\241\000\001',
	WordMsoTriStateToggle = '\000\240\377\375',
	WordMsoTriStateTriStateUnset = '\000\240\377\376'
};
typedef enum WordMsoTriState WordMsoTriState;

enum WordMsoBlackWhiteMode {
	WordMsoBlackWhiteModeBlackAndWhiteUnset = '\000\254\377\376',
	WordMsoBlackWhiteModeBlackAndWhiteModeAutomatic = '\000\255\000\001',
	WordMsoBlackWhiteModeBlackAndWhiteModeGrayScale = '\000\255\000\002',
	WordMsoBlackWhiteModeBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	WordMsoBlackWhiteModeBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	WordMsoBlackWhiteModeBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	WordMsoBlackWhiteModeBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	WordMsoBlackWhiteModeBlackAndWhiteModeHighContrast = '\000\255\000\007',
	WordMsoBlackWhiteModeBlackAndWhiteModeBlack = '\000\255\000\010',
	WordMsoBlackWhiteModeBlackAndWhiteModeWhite = '\000\255\000\011',
	WordMsoBlackWhiteModeBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum WordMsoBlackWhiteMode WordMsoBlackWhiteMode;

enum WordMsoBarPosition {
	WordMsoBarPositionBarLeft = '\000r\000\000',
	WordMsoBarPositionBarTop = '\000r\000\001',
	WordMsoBarPositionBarRight = '\000r\000\002',
	WordMsoBarPositionBarBottom = '\000r\000\003',
	WordMsoBarPositionBarFloating = '\000r\000\004',
	WordMsoBarPositionBarPopUp = '\000r\000\005',
	WordMsoBarPositionBarMenu = '\000r\000\006'
};
typedef enum WordMsoBarPosition WordMsoBarPosition;

enum WordMsoBarProtection {
	WordMsoBarProtectionNoProtection = '\000s\000\000',
	WordMsoBarProtectionNoCustomize = '\000s\000\001',
	WordMsoBarProtectionNoResize = '\000s\000\002',
	WordMsoBarProtectionNoMove = '\000s\000\004',
	WordMsoBarProtectionNoChangeVisible = '\000s\000\010',
	WordMsoBarProtectionNoChangeDock = '\000s\000\020',
	WordMsoBarProtectionNoVerticalDock = '\000s\000 ',
	WordMsoBarProtectionNoHorizontalDock = '\000s\000@'
};
typedef enum WordMsoBarProtection WordMsoBarProtection;

enum WordMsoBarType {
	WordMsoBarTypeNormalCommandBar = '\000t\000\000',
	WordMsoBarTypeMenubarCommandBar = '\000t\000\001',
	WordMsoBarTypePopupCommandBar = '\000t\000\002'
};
typedef enum WordMsoBarType WordMsoBarType;

enum WordMsoControlType {
	WordMsoControlTypeControlCustom = '\000u\000\000',
	WordMsoControlTypeControlButton = '\000u\000\001',
	WordMsoControlTypeControlEdit = '\000u\000\002',
	WordMsoControlTypeControlDropDown = '\000u\000\003',
	WordMsoControlTypeControlCombobox = '\000u\000\004',
	WordMsoControlTypeButtonDropDown = '\000u\000\005',
	WordMsoControlTypeSplitDropDown = '\000u\000\006',
	WordMsoControlTypeOCXDropDown = '\000u\000\007',
	WordMsoControlTypeGenericDropDown = '\000u\000\010',
	WordMsoControlTypeGraphicDropDown = '\000u\000\011',
	WordMsoControlTypeControlPopup = '\000u\000\012',
	WordMsoControlTypeGraphicPopup = '\000u\000\013',
	WordMsoControlTypeButtonPopup = '\000u\000\014',
	WordMsoControlTypeSplitButtonPopup = '\000u\000\015',
	WordMsoControlTypeSplitButtonMRUPopup = '\000u\000\016',
	WordMsoControlTypeControlLabel = '\000u\000\017',
	WordMsoControlTypeExpandingGrid = '\000u\000\020',
	WordMsoControlTypeSplitExpandingGrid = '\000u\000\021',
	WordMsoControlTypeControlGrid = '\000u\000\022',
	WordMsoControlTypeControlGauge = '\000u\000\023',
	WordMsoControlTypeGraphicCombobox = '\000u\000\024',
	WordMsoControlTypeControlPane = '\000u\000\025',
	WordMsoControlTypeActiveX = '\000u\000\026',
	WordMsoControlTypeControlGroup = '\000u\000\027',
	WordMsoControlTypeControlTab = '\000u\000\030',
	WordMsoControlTypeControlSpinner = '\000u\000\031'
};
typedef enum WordMsoControlType WordMsoControlType;

enum WordMsoButtonState {
	WordMsoButtonStateButtonStateUp = '\000v\000\000',
	WordMsoButtonStateButtonStateDown = '\000u\377\377',
	WordMsoButtonStateButtonStateUnset = '\000v\000\002'
};
typedef enum WordMsoButtonState WordMsoButtonState;

enum WordMsoControlOLEUsage {
	WordMsoControlOLEUsageNeither = '\000w\000\000',
	WordMsoControlOLEUsageServer = '\000w\000\001',
	WordMsoControlOLEUsageClient = '\000w\000\002',
	WordMsoControlOLEUsageBoth = '\000w\000\003'
};
typedef enum WordMsoControlOLEUsage WordMsoControlOLEUsage;

enum WordMsoButtonStyle {
	WordMsoButtonStyleButtonAutomatic = '\000x\000\000',
	WordMsoButtonStyleButtonIcon = '\000x\000\001',
	WordMsoButtonStyleButtonCaption = '\000x\000\002',
	WordMsoButtonStyleButtonIconAndCaption = '\000x\000\003'
};
typedef enum WordMsoButtonStyle WordMsoButtonStyle;

enum WordMsoComboStyle {
	WordMsoComboStyleComboboxStyleNormal = '\000y\000\000',
	WordMsoComboStyleComboboxStyleLabel = '\000y\000\001'
};
typedef enum WordMsoComboStyle WordMsoComboStyle;

enum WordMsoMenuAnimation {
	WordMsoMenuAnimationNone = '\000{\000\000',
	WordMsoMenuAnimationRandom = '\000{\000\001',
	WordMsoMenuAnimationUnfold = '\000{\000\002',
	WordMsoMenuAnimationSlide = '\000{\000\003'
};
typedef enum WordMsoMenuAnimation WordMsoMenuAnimation;

enum WordMsoHyperlinkType {
	WordMsoHyperlinkTypeHyperlinkTypeTextRange = '\000\226\000\000',
	WordMsoHyperlinkTypeHyperlinkTypeShape = '\000\226\000\001',
	WordMsoHyperlinkTypeHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum WordMsoHyperlinkType WordMsoHyperlinkType;

enum WordMsoExtraInfoMethod {
	WordMsoExtraInfoMethodAppendString = '\000\256\000\000',
	WordMsoExtraInfoMethodPostString = '\000\256\000\001'
};
typedef enum WordMsoExtraInfoMethod WordMsoExtraInfoMethod;

enum WordMsoAnimationType {
	WordMsoAnimationTypeIdle = '\000|\000\001',
	WordMsoAnimationTypeGreeting = '\000|\000\002',
	WordMsoAnimationTypeGoodbye = '\000|\000\003',
	WordMsoAnimationTypeBeginSpeaking = '\000|\000\004',
	WordMsoAnimationTypeCharacterSuccessMajor = '\000|\000\006',
	WordMsoAnimationTypeGetAttentionMajor = '\000|\000\013',
	WordMsoAnimationTypeGetAttentionMinor = '\000|\000\014',
	WordMsoAnimationTypeSearching = '\000|\000\015',
	WordMsoAnimationTypePrinting = '\000|\000\022',
	WordMsoAnimationTypeGestureRight = '\000|\000\023',
	WordMsoAnimationTypeWritingNotingSomething = '\000|\000\026',
	WordMsoAnimationTypeWorkingAtSomething = '\000|\000\027',
	WordMsoAnimationTypeThinking = '\000|\000\030',
	WordMsoAnimationTypeSendingMail = '\000|\000\031',
	WordMsoAnimationTypeListensToComputer = '\000|\000\032',
	WordMsoAnimationTypeDisappear = '\000|\000\037',
	WordMsoAnimationTypeAppear = '\000|\000 ',
	WordMsoAnimationTypeGetArtsy = '\000|\000d',
	WordMsoAnimationTypeGetTechy = '\000|\000e',
	WordMsoAnimationTypeGetWizardy = '\000|\000f',
	WordMsoAnimationTypeCheckingSomething = '\000|\000g',
	WordMsoAnimationTypeLookDown = '\000|\000h',
	WordMsoAnimationTypeLookDownLeft = '\000|\000i',
	WordMsoAnimationTypeLookDownRight = '\000|\000j',
	WordMsoAnimationTypeLookLeft = '\000|\000k',
	WordMsoAnimationTypeLookRight = '\000|\000l',
	WordMsoAnimationTypeLookUp = '\000|\000m',
	WordMsoAnimationTypeLookUpLeft = '\000|\000n',
	WordMsoAnimationTypeLookUpRight = '\000|\000o',
	WordMsoAnimationTypeSaving = '\000|\000p',
	WordMsoAnimationTypeGestureDown = '\000|\000q',
	WordMsoAnimationTypeGestureLeft = '\000|\000r',
	WordMsoAnimationTypeGestureUp = '\000|\000s',
	WordMsoAnimationTypeEmptyTrash = '\000|\000t'
};
typedef enum WordMsoAnimationType WordMsoAnimationType;

enum WordMsoButtonSetType {
	WordMsoButtonSetTypeButtonNone = '\000}\000\000',
	WordMsoButtonSetTypeButtonOk = '\000}\000\001',
	WordMsoButtonSetTypeButtonCancel = '\000}\000\002',
	WordMsoButtonSetTypeButtonsOkCancel = '\000}\000\003',
	WordMsoButtonSetTypeButtonsYesNo = '\000}\000\004',
	WordMsoButtonSetTypeButtonsYesNoCancel = '\000}\000\005',
	WordMsoButtonSetTypeButtonsBackClose = '\000}\000\006',
	WordMsoButtonSetTypeButtonsNextClose = '\000}\000\007',
	WordMsoButtonSetTypeButtonsBackNextClose = '\000}\000\010',
	WordMsoButtonSetTypeButtonsRetryCancel = '\000}\000\011',
	WordMsoButtonSetTypeButtonsAbortRetryIgnore = '\000}\000\012',
	WordMsoButtonSetTypeButtonsSearchClose = '\000}\000\013',
	WordMsoButtonSetTypeButtonsBackNextSnooze = '\000}\000\014',
	WordMsoButtonSetTypeButtonsTipsOptionsClose = '\000}\000\015',
	WordMsoButtonSetTypeButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum WordMsoButtonSetType WordMsoButtonSetType;

enum WordMsoIconType {
	WordMsoIconTypeIconNone = '\000~\000\000',
	WordMsoIconTypeIconApplication = '\000~\000\001',
	WordMsoIconTypeIconAlert = '\000~\000\002',
	WordMsoIconTypeIconTip = '\000~\000\003',
	WordMsoIconTypeIconAlertCritical = '\000~\000e',
	WordMsoIconTypeIconAlertWarning = '\000~\000g',
	WordMsoIconTypeIconAlertInfo = '\000~\000h'
};
typedef enum WordMsoIconType WordMsoIconType;

enum WordMsoWizardActType {
	WordMsoWizardActTypeInactive = '\000\202\000\000',
	WordMsoWizardActTypeActive = '\000\202\000\001',
	WordMsoWizardActTypeSuspend = '\000\202\000\002',
	WordMsoWizardActTypeResume = '\000\202\000\003'
};
typedef enum WordMsoWizardActType WordMsoWizardActType;

enum WordMsoDocProperties {
	WordMsoDocPropertiesPropertyTypeNumber = '\000\242\000\001',
	WordMsoDocPropertiesPropertyTypeBoolean = '\000\242\000\002',
	WordMsoDocPropertiesPropertyTypeDate = '\000\242\000\003',
	WordMsoDocPropertiesPropertyTypeString = '\000\242\000\004',
	WordMsoDocPropertiesPropertyTypeFloat = '\000\242\000\005'
};
typedef enum WordMsoDocProperties WordMsoDocProperties;

enum WordMsoAutomationSecurity {
	WordMsoAutomationSecurityMsoAutomationSecurityLow = '\000\243\000\001',
	WordMsoAutomationSecurityMsoAutomationSecurityByUI = '\000\243\000\002',
	WordMsoAutomationSecurityMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum WordMsoAutomationSecurity WordMsoAutomationSecurity;

enum WordMsoScreenSize {
	WordMsoScreenSizeResolution544x376 = '\000\204\000\000',
	WordMsoScreenSizeResolution640x480 = '\000\204\000\001',
	WordMsoScreenSizeResolution720x512 = '\000\204\000\002',
	WordMsoScreenSizeResolution800x600 = '\000\204\000\003',
	WordMsoScreenSizeResolution1024x768 = '\000\204\000\004',
	WordMsoScreenSizeResolution1152x882 = '\000\204\000\005',
	WordMsoScreenSizeResolution1152x900 = '\000\204\000\006',
	WordMsoScreenSizeResolution1280x1024 = '\000\204\000\007',
	WordMsoScreenSizeResolution1600x1200 = '\000\204\000\010',
	WordMsoScreenSizeResolution1800x1440 = '\000\204\000\011',
	WordMsoScreenSizeResolution1920x1200 = '\000\204\000\012'
};
typedef enum WordMsoScreenSize WordMsoScreenSize;

enum WordMsoCharacterSet {
	WordMsoCharacterSetArabicCharacterSet = '\000\205\000\001',
	WordMsoCharacterSetCyrillicCharacterSet = '\000\205\000\002',
	WordMsoCharacterSetEnglishCharacterSet = '\000\205\000\003',
	WordMsoCharacterSetGreekCharacterSet = '\000\205\000\004',
	WordMsoCharacterSetHebrewCharacterSet = '\000\205\000\005',
	WordMsoCharacterSetJapaneseCharacterSet = '\000\205\000\006',
	WordMsoCharacterSetKoreanCharacterSet = '\000\205\000\007',
	WordMsoCharacterSetMultilingualUnicodeCharacterSet = '\000\205\000\010',
	WordMsoCharacterSetSimplifiedChineseCharacterSet = '\000\205\000\011',
	WordMsoCharacterSetThaiCharacterSet = '\000\205\000\012',
	WordMsoCharacterSetTraditionalChineseCharacterSet = '\000\205\000\013',
	WordMsoCharacterSetVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum WordMsoCharacterSet WordMsoCharacterSet;

enum WordMsoEncoding {
	WordMsoEncodingEncodingThai = '\000\213\003j',
	WordMsoEncodingEncodingJapaneseShiftJIS = '\000\213\003\244',
	WordMsoEncodingEncodingSimplifiedChinese = '\000\213\003\250',
	WordMsoEncodingEncodingKorean = '\000\213\003\265',
	WordMsoEncodingEncodingBig5TraditionalChinese = '\000\213\003\266',
	WordMsoEncodingEncodingLittleEndian = '\000\213\004\260',
	WordMsoEncodingEncodingBigEndian = '\000\213\004\261',
	WordMsoEncodingEncodingCentralEuropean = '\000\213\004\342',
	WordMsoEncodingEncodingCyrillic = '\000\213\004\343',
	WordMsoEncodingEncodingWestern = '\000\213\004\344',
	WordMsoEncodingEncodingGreek = '\000\213\004\345',
	WordMsoEncodingEncodingTurkish = '\000\213\004\346',
	WordMsoEncodingEncodingHebrew = '\000\213\004\347',
	WordMsoEncodingEncodingArabic = '\000\213\004\350',
	WordMsoEncodingEncodingBaltic = '\000\213\004\351',
	WordMsoEncodingEncodingVietnamese = '\000\213\004\352',
	WordMsoEncodingEncodingISO88591Latin1 = '\000\213o\257',
	WordMsoEncodingEncodingISO88592CentralEurope = '\000\213o\260',
	WordMsoEncodingEncodingISO88593Latin3 = '\000\213o\261',
	WordMsoEncodingEncodingISO88594Baltic = '\000\213o\262',
	WordMsoEncodingEncodingISO88595Cyrillic = '\000\213o\263',
	WordMsoEncodingEncodingISO88596Arabic = '\000\213o\264',
	WordMsoEncodingEncodingISO88597Greek = '\000\213o\265',
	WordMsoEncodingEncodingISO88598Hebrew = '\000\213o\266',
	WordMsoEncodingEncodingISO88599Turkish = '\000\213o\267',
	WordMsoEncodingEncodingISO885915Latin9 = '\000\213o\275',
	WordMsoEncodingEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	WordMsoEncodingEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	WordMsoEncodingEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	WordMsoEncodingEncodingISO2022KR = '\000\213\3041',
	WordMsoEncodingEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	WordMsoEncodingEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	WordMsoEncodingEncodingMacRoman = '\000\213\'\020',
	WordMsoEncodingEncodingMacJapanese = '\000\213\'\021',
	WordMsoEncodingEncodingMacTraditionalChinese = '\000\213\'\022',
	WordMsoEncodingEncodingMacKorean = '\000\213\'\023',
	WordMsoEncodingEncodingMacArabic = '\000\213\'\024',
	WordMsoEncodingEncodingMacHebrew = '\000\213\'\025',
	WordMsoEncodingEncodingMacGreek1 = '\000\213\'\026',
	WordMsoEncodingEncodingMacCyrillic = '\000\213\'\027',
	WordMsoEncodingEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	WordMsoEncodingEncodingMacRomania = '\000\213\'\032',
	WordMsoEncodingEncodingMacUkraine = '\000\213\'!',
	WordMsoEncodingEncodingMacLatin2 = '\000\213\'-',
	WordMsoEncodingEncodingMacIcelandic = '\000\213\'_',
	WordMsoEncodingEncodingMacTurkish = '\000\213\'a',
	WordMsoEncodingEncodingMacCroatia = '\000\213\'b',
	WordMsoEncodingEncodingEBCDICUSCanada = '\000\213\000%',
	WordMsoEncodingEncodingEBCDICInternational = '\000\213\001\364',
	WordMsoEncodingEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	WordMsoEncodingEncodingEBCDICGreekModern = '\000\213\003k',
	WordMsoEncodingEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	WordMsoEncodingEncodingEBCDICGermany = '\000\213O1',
	WordMsoEncodingEncodingEBCDICDenmarkNorway = '\000\213O5',
	WordMsoEncodingEncodingEBCDICFinlandSweden = '\000\213O6',
	WordMsoEncodingEncodingEBCDICItaly = '\000\213O8',
	WordMsoEncodingEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	WordMsoEncodingEncodingEBCDICUnitedKingdom = '\000\213O=',
	WordMsoEncodingEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	WordMsoEncodingEncodingEBCDICFrance = '\000\213OI',
	WordMsoEncodingEncodingEBCDICArabic = '\000\213O\304',
	WordMsoEncodingEncodingEBCDICGreek = '\000\213O\307',
	WordMsoEncodingEncodingEBCDICHebrew = '\000\213O\310',
	WordMsoEncodingEncodingEBCDICKoreanExtended = '\000\213Qa',
	WordMsoEncodingEncodingEBCDICThai = '\000\213Qf',
	WordMsoEncodingEncodingEBCDICIcelandic = '\000\213Q\207',
	WordMsoEncodingEncodingEBCDICTurkish = '\000\213Q\251',
	WordMsoEncodingEncodingEBCDICRussian = '\000\213Q\220',
	WordMsoEncodingEncodingEBCDICSerbianBulgarian = '\000\213R!',
	WordMsoEncodingEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	WordMsoEncodingEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	WordMsoEncodingEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	WordMsoEncodingEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	WordMsoEncodingEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	WordMsoEncodingEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	WordMsoEncodingEncodingOEMUnitedStates = '\000\213\001\265',
	WordMsoEncodingEncodingOEMGreek = '\000\213\002\341',
	WordMsoEncodingEncodingOEMBaltic = '\000\213\003\007',
	WordMsoEncodingEncodingOEMMultilingualLatinI = '\000\213\003R',
	WordMsoEncodingEncodingOEMMultilingualLatinII = '\000\213\003T',
	WordMsoEncodingEncodingOEMCyrillic = '\000\213\003W',
	WordMsoEncodingEncodingOEMTurkish = '\000\213\003Y',
	WordMsoEncodingEncodingOEMPortuguese = '\000\213\003\\',
	WordMsoEncodingEncodingOEMIcelandic = '\000\213\003]',
	WordMsoEncodingEncodingOEMHebrew = '\000\213\003^',
	WordMsoEncodingEncodingOEMCanadianFrench = '\000\213\003_',
	WordMsoEncodingEncodingOEMArabic = '\000\213\003`',
	WordMsoEncodingEncodingOEMNordic = '\000\213\003a',
	WordMsoEncodingEncodingOEMCyrillicII = '\000\213\003b',
	WordMsoEncodingEncodingOEMModernGreek = '\000\213\003e',
	WordMsoEncodingEncodingEUCJapanese = '\000\213\312\334',
	WordMsoEncodingEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	WordMsoEncodingEncodingEUCKorean = '\000\213\312\355',
	WordMsoEncodingEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	WordMsoEncodingEncodingDevanagari = '\000\213\336\252',
	WordMsoEncodingEncodingBengali = '\000\213\336\253',
	WordMsoEncodingEncodingTamil = '\000\213\336\254',
	WordMsoEncodingEncodingTelugu = '\000\213\336\255',
	WordMsoEncodingEncodingAssamese = '\000\213\336\256',
	WordMsoEncodingEncodingOriya = '\000\213\336\257',
	WordMsoEncodingEncodingKannada = '\000\213\336\260',
	WordMsoEncodingEncodingMalayalam = '\000\213\336\261',
	WordMsoEncodingEncodingGujarati = '\000\213\336\262',
	WordMsoEncodingEncodingPunjabi = '\000\213\336\263',
	WordMsoEncodingEncodingArabicASMO = '\000\213\002\304',
	WordMsoEncodingEncodingArabicTransparentASMO = '\000\213\002\320',
	WordMsoEncodingEncodingKoreanJohab = '\000\213\005Q',
	WordMsoEncodingEncodingTaiwanCNS = '\000\213N ',
	WordMsoEncodingEncodingTaiwanTCA = '\000\213N!',
	WordMsoEncodingEncodingTaiwanEten = '\000\213N\"',
	WordMsoEncodingEncodingTaiwanIBM5550 = '\000\213N#',
	WordMsoEncodingEncodingTaiwanTeletext = '\000\213N$',
	WordMsoEncodingEncodingTaiwanWang = '\000\213N%',
	WordMsoEncodingEncodingIA5IRV = '\000\213N\211',
	WordMsoEncodingEncodingIA5German = '\000\213N\212',
	WordMsoEncodingEncodingIA5Swedish = '\000\213N\213',
	WordMsoEncodingEncodingIA5Norwegian = '\000\213N\214',
	WordMsoEncodingEncodingUSASCII = '\000\213N\237',
	WordMsoEncodingEncodingT61 = '\000\213O%',
	WordMsoEncodingEncodingISO6937NonspacingAccent = '\000\213O-',
	WordMsoEncodingEncodingKOI8R = '\000\213Q\202',
	WordMsoEncodingEncodingExtAlphaLowercase = '\000\213R#',
	WordMsoEncodingEncodingKOI8U = '\000\213Uj',
	WordMsoEncodingEncodingEuropa3 = '\000\213qI',
	WordMsoEncodingEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	WordMsoEncodingEncodingUTF7 = '\000\213\375\350',
	WordMsoEncodingEncodingUTF8 = '\000\213\375\351'
};
typedef enum WordMsoEncoding WordMsoEncoding;

enum WordReset {
	WordResetCommandBar = 'msCB',
	WordResetCommandBarControl = 'mCBC'
};
typedef enum WordReset WordReset;

enum WordWdMoveSelection {
	WordWdMoveSelectionTowardsTheEnd = '\002\266\000\000',
	WordWdMoveSelectionTowardsTheStart = '\002\266\000\001',
	WordWdMoveSelectionTowardsTheLeft = '\002\266\000\002',
	WordWdMoveSelectionTowardsTheRight = '\002\266\000\003',
	WordWdMoveSelectionTowardsTheTop = '\002\266\000\004',
	WordWdMoveSelectionTowardsTheBottom = '\002\266\000\005'
};
typedef enum WordWdMoveSelection WordWdMoveSelection;

enum WordWdMoveUntilWhile {
	WordWdMoveUntilWhileTowardsTheTail = '\002\267\000\000',
	WordWdMoveUntilWhileTowardsTheBeginning = '\002\267\000\001'
};
typedef enum WordWdMoveUntilWhile WordWdMoveUntilWhile;

enum WordWdInsertAboveBelow {
	WordWdInsertAboveBelowAbove = '\002\270\000\000',
	WordWdInsertAboveBelowBelow = '\002\270\000\001'
};
typedef enum WordWdInsertAboveBelow WordWdInsertAboveBelow;

enum WordWdInsertBeforeAfter {
	WordWdInsertBeforeAfterBeforeTheObject = '\002\271\000\000',
	WordWdInsertBeforeAfterAfterTheObject = '\002\271\000\001'
};
typedef enum WordWdInsertBeforeAfter WordWdInsertBeforeAfter;

enum WordWdInsertRightLeft {
	WordWdInsertRightLeftInsertOnTheRight = '\002\273\000\000',
	WordWdInsertRightLeftInsertOnTheLeft = '\002\273\000\001'
};
typedef enum WordWdInsertRightLeft WordWdInsertRightLeft;

enum WordWdInsertionPoint {
	WordWdInsertionPointOffset = '\003\037\000\000',
	WordWdInsertionPointLocationReference = '\003\037\000\001'
};
typedef enum WordWdInsertionPoint WordWdInsertionPoint;

enum WordWdTemplateType {
	WordWdTemplateTypeTypeNormalTemplate = '\001\365\000\000',
	WordWdTemplateTypeTypeGlobalTemplate = '\001\365\000\001',
	WordWdTemplateTypeTypeAttachedTemplate = '\001\365\000\002'
};
typedef enum WordWdTemplateType WordWdTemplateType;

enum WordWdContinue {
	WordWdContinueContinueDisabled = '\001\366\000\000',
	WordWdContinueResetList = '\001\366\000\001',
	WordWdContinueContinueList = '\001\366\000\002'
};
typedef enum WordWdContinue WordWdContinue;

enum WordWdIMEMode {
	WordWdIMEModeIMEModeNoControl = '\001\367\000\000',
	WordWdIMEModeIMEModeOn = '\001\367\000\001',
	WordWdIMEModeIMEModeOff = '\001\367\000\002',
	WordWdIMEModeIMEModeHiragana = '\001\367\000\004',
	WordWdIMEModeIMEModeKatakana = '\001\367\000\005',
	WordWdIMEModeIMEModeKatakanaHalf = '\001\367\000\006',
	WordWdIMEModeIMEModeAlphaFull = '\001\367\000\007',
	WordWdIMEModeIMEModeAlpha = '\001\367\000\010',
	WordWdIMEModeIMEModeHangulFull = '\001\367\000\011',
	WordWdIMEModeIMEModeHangul = '\001\367\000\012'
};
typedef enum WordWdIMEMode WordWdIMEMode;

enum WordWdBaselineAlignment {
	WordWdBaselineAlignmentBaselineAlignTop = '\001\370\000\000',
	WordWdBaselineAlignmentBaselineAlignCenter = '\001\370\000\001',
	WordWdBaselineAlignmentBaselineAlignBaseline = '\001\370\000\002',
	WordWdBaselineAlignmentBaselineAlignEastAsian50 = '\001\370\000\003',
	WordWdBaselineAlignmentBaselineAlignAuto = '\001\370\000\004'
};
typedef enum WordWdBaselineAlignment WordWdBaselineAlignment;

enum WordWdIndexFilter {
	WordWdIndexFilterIndexFilterNone = '\001\371\000\000',
	WordWdIndexFilterIndexFilterAiueo = '\001\371\000\001',
	WordWdIndexFilterIndexFilterAkasatana = '\001\371\000\002',
	WordWdIndexFilterIndexFilterChosung = '\001\371\000\003',
	WordWdIndexFilterIndexFilterLow = '\001\371\000\004',
	WordWdIndexFilterIndexFilterMedium = '\001\371\000\005',
	WordWdIndexFilterIndexFilterFull = '\001\371\000\006'
};
typedef enum WordWdIndexFilter WordWdIndexFilter;

enum WordWdIndexSortBy {
	WordWdIndexSortByIndexSortByStroke = '\001\372\000\000',
	WordWdIndexSortByIndexSortBySyllable = '\001\372\000\001'
};
typedef enum WordWdIndexSortBy WordWdIndexSortBy;

enum WordWdJustificationMode {
	WordWdJustificationModeJustificationModeExpand = '\001\373\000\000',
	WordWdJustificationModeJustificationModeCompress = '\001\373\000\001',
	WordWdJustificationModeJustificationModeCompressKana = '\001\373\000\002'
};
typedef enum WordWdJustificationMode WordWdJustificationMode;

enum WordWdFarEastLineBreakLevel {
	WordWdFarEastLineBreakLevelEastAsianLineBreakLevelNormal = '\001\374\000\000',
	WordWdFarEastLineBreakLevelEastAsianLineBreakLevelStrict = '\001\374\000\001',
	WordWdFarEastLineBreakLevelEastAsianLineBreakLevelCustom = '\001\374\000\002'
};
typedef enum WordWdFarEastLineBreakLevel WordWdFarEastLineBreakLevel;

enum WordWdMultipleWordConversionsMode {
	WordWdMultipleWordConversionsModeHangulToHanja = '\001\375\000\000',
	WordWdMultipleWordConversionsModeHanjaToHangul = '\001\375\000\001'
};
typedef enum WordWdMultipleWordConversionsMode WordWdMultipleWordConversionsMode;

enum WordWdColorIndex {
	WordWdColorIndexAuto = '\001\376\000\000',
	WordWdColorIndexBlack = '\001\376\000\001',
	WordWdColorIndexBlue = '\001\376\000\002',
	WordWdColorIndexTurquoise = '\001\376\000\003',
	WordWdColorIndexBrightGreen = '\001\376\000\004',
	WordWdColorIndexPink = '\001\376\000\005',
	WordWdColorIndexRed = '\001\376\000\006',
	WordWdColorIndexYellow = '\001\376\000\007',
	WordWdColorIndexWhite = '\001\376\000\010',
	WordWdColorIndexDarkBlue = '\001\376\000\011',
	WordWdColorIndexTeal = '\001\376\000\012',
	WordWdColorIndexGreen = '\001\376\000\013',
	WordWdColorIndexViolet = '\001\376\000\014',
	WordWdColorIndexDarkRed = '\001\376\000\015',
	WordWdColorIndexDarkYellow = '\001\376\000\016',
	WordWdColorIndexGray50 = '\001\376\000\017',
	WordWdColorIndexGray25 = '\001\376\000\020',
	WordWdColorIndexByAuthor = '\001\375\377\377',
	WordWdColorIndexNoHighlight = '\001\376\000\000'
};
typedef enum WordWdColorIndex WordWdColorIndex;

enum WordWdColor {
	WordWdColorColorAutomatic = '\377\000\000\000',
	WordWdColorColorBlack = '\000\000\000\000',
	WordWdColorColorBlue = '\000\377\000\000',
	WordWdColorColorPink = '\000\377\000\377',
	WordWdColorColorRed = '\000\000\000\377',
	WordWdColorColorYellow = '\000\000\377\377',
	WordWdColorColorTurquoise = '\000\377\377\000',
	WordWdColorColorBrightGreen = '\000\000\377\000',
	WordWdColorColorGreen = '\000\000\200\000',
	WordWdColorColorWhite = '\000\377\377\377',
	WordWdColorColorDarkBlue = '\000\200\000\000',
	WordWdColorColorTeal = '\000\200\200\000',
	WordWdColorColorViolet = '\000\200\000\200',
	WordWdColorColorDarkGreen = '\000\0003\000',
	WordWdColorColorDarkRed = '\000\000\000\200',
	WordWdColorColorDarkYellow = '\000\000\200\200',
	WordWdColorColorBrown = '\000\0003\231',
	WordWdColorColorOliveGreen = '\000\00033',
	WordWdColorColorDarkTeal = '\000f3\000',
	WordWdColorColorIndigo = '\000\23133',
	WordWdColorColorOrange = '\000\000f\377',
	WordWdColorColorBlueGray = '\000\231ff',
	WordWdColorColorLightOrange = '\000\000\231\377',
	WordWdColorColorLime = '\000\000\314\231',
	WordWdColorColorSeaGreen = '\000f\2313',
	WordWdColorColorAqua = '\000\314\3143',
	WordWdColorColorLightBlue = '\000\377f3',
	WordWdColorColorGold = '\000\000\314\377',
	WordWdColorColorSkyBlue = '\000\377\314\000',
	WordWdColorColorPlum = '\000f3\231',
	WordWdColorColorRose = '\000\314\231\377',
	WordWdColorColorTan = '\000\231\314\377',
	WordWdColorColorLightYellow = '\000\231\377\377',
	WordWdColorColorLightGreen = '\000\314\377\314',
	WordWdColorColorLightTurquoise = '\000\377\377\314',
	WordWdColorColorPaleBlue = '\000\377\314\231',
	WordWdColorColorLavender = '\000\377\231\314',
	WordWdColorColorGray05 = '\000\363\363\363',
	WordWdColorColorGray10 = '\000\346\346\346',
	WordWdColorColorGray125 = '\000\340\340\340',
	WordWdColorColorGray15 = '\000\331\331\331',
	WordWdColorColorGray20 = '\000\314\314\314',
	WordWdColorColorGray25 = '\000\300\300\300',
	WordWdColorColorGray30 = '\000\263\263\263',
	WordWdColorColorGray35 = '\000\246\246\246',
	WordWdColorColorGray375 = '\000\240\240\240',
	WordWdColorColorGray40 = '\000\231\231\231',
	WordWdColorColorGray45 = '\000\214\214\214',
	WordWdColorColorGray50 = '\000\200\200\200',
	WordWdColorColorGray55 = '\000sss',
	WordWdColorColorGray60 = '\000fff',
	WordWdColorColorGray625 = '\000```',
	WordWdColorColorGray65 = '\000YYY',
	WordWdColorColorGray70 = '\000LLL',
	WordWdColorColorGray75 = '\000@@@',
	WordWdColorColorGray80 = '\000333',
	WordWdColorColorGray85 = '\000&&&',
	WordWdColorColorGray875 = '\000   ',
	WordWdColorColorGray90 = '\000\031\031\031',
	WordWdColorColorGray95 = '\000\014\014\014'
};
typedef enum WordWdColor WordWdColor;

enum WordWdTextureIndex {
	WordWdTextureIndexTextureNone = '\002\000\000\000',
	WordWdTextureIndexTexture2Pt5Percent = '\002\000\000\031',
	WordWdTextureIndexTexture5Percent = '\002\000\0002',
	WordWdTextureIndexTexture7Pt5Percent = '\002\000\000K',
	WordWdTextureIndexTexture10Percent = '\002\000\000d',
	WordWdTextureIndexTexture12Pt5Percent = '\002\000\000}',
	WordWdTextureIndexTexture15Percent = '\002\000\000\226',
	WordWdTextureIndexTexture17Pt5Percent = '\002\000\000\257',
	WordWdTextureIndexTexture20Percent = '\002\000\000\310',
	WordWdTextureIndexTexture22Pt5Percent = '\002\000\000\341',
	WordWdTextureIndexTexture25Percent = '\002\000\000\372',
	WordWdTextureIndexTexture27Pt5Percent = '\002\000\001\023',
	WordWdTextureIndexTexture30Percent = '\002\000\001,',
	WordWdTextureIndexTexture32Pt5Percent = '\002\000\001E',
	WordWdTextureIndexTexture35Percent = '\002\000\001^',
	WordWdTextureIndexTexture37Pt5Percent = '\002\000\001w',
	WordWdTextureIndexTexture40Percent = '\002\000\001\220',
	WordWdTextureIndexTexture42Pt5Percent = '\002\000\001\251',
	WordWdTextureIndexTexture45Percent = '\002\000\001\302',
	WordWdTextureIndexTexture47Pt5Percent = '\002\000\001\333',
	WordWdTextureIndexTexture50Percent = '\002\000\001\364',
	WordWdTextureIndexTexture52Pt5Percent = '\002\000\002\015',
	WordWdTextureIndexTexture55Percent = '\002\000\002&',
	WordWdTextureIndexTexture57Pt5Percent = '\002\000\002\?',
	WordWdTextureIndexTexture60Percent = '\002\000\002X',
	WordWdTextureIndexTexture62Pt5Percent = '\002\000\002q',
	WordWdTextureIndexTexture65Percent = '\002\000\002\212',
	WordWdTextureIndexTexture67Pt5Percent = '\002\000\002\243',
	WordWdTextureIndexTexture70Percent = '\002\000\002\274',
	WordWdTextureIndexTexture72Pt5Percent = '\002\000\002\325',
	WordWdTextureIndexTexture75Percent = '\002\000\002\356',
	WordWdTextureIndexTexture77Pt5Percent = '\002\000\003\007',
	WordWdTextureIndexTexture80Percent = '\002\000\003 ',
	WordWdTextureIndexTexture82Pt5Percent = '\002\000\0039',
	WordWdTextureIndexTexture85Percent = '\002\000\003R',
	WordWdTextureIndexTexture87Pt5Percent = '\002\000\003k',
	WordWdTextureIndexTexture90Percent = '\002\000\003\204',
	WordWdTextureIndexTexture92Pt5Percent = '\002\000\003\235',
	WordWdTextureIndexTexture95Percent = '\002\000\003\266',
	WordWdTextureIndexTexture97Pt5Percent = '\002\000\003\317',
	WordWdTextureIndexTextureSolid = '\002\000\003\350',
	WordWdTextureIndexTextureDarkHorizontal = '\001\377\377\377',
	WordWdTextureIndexTextureDarkVertical = '\001\377\377\376',
	WordWdTextureIndexTextureDarkDiagonalDown = '\001\377\377\375',
	WordWdTextureIndexTextureDarkDiagonalUp = '\001\377\377\374',
	WordWdTextureIndexTextureDarkCross = '\001\377\377\373',
	WordWdTextureIndexTextureDarkDiagonalCross = '\001\377\377\372',
	WordWdTextureIndexTextureHorizontal = '\001\377\377\371',
	WordWdTextureIndexTextureVertical = '\001\377\377\370',
	WordWdTextureIndexTextureDiagonalDown = '\001\377\377\367',
	WordWdTextureIndexTextureDiagonalUp = '\001\377\377\366',
	WordWdTextureIndexTextureCross = '\001\377\377\365',
	WordWdTextureIndexTextureDiagonalCross = '\001\377\377\364'
};
typedef enum WordWdTextureIndex WordWdTextureIndex;

enum WordWdUnderline {
	WordWdUnderlineUnderlineNone = '\002\001\000\000',
	WordWdUnderlineUnderlineSingle = '\002\001\000\001',
	WordWdUnderlineUnderlineWords = '\002\001\000\002',
	WordWdUnderlineUnderlineDouble = '\002\001\000\003',
	WordWdUnderlineUnderlineDotted = '\002\001\000\004',
	WordWdUnderlineUnderlineThick = '\002\001\000\006',
	WordWdUnderlineUnderlineDash = '\002\001\000\007',
	WordWdUnderlineUnderlineDotDash = '\002\001\000\011',
	WordWdUnderlineUnderlineDotDotDash = '\002\001\000\012',
	WordWdUnderlineUnderlineWavy = '\002\001\000\013',
	WordWdUnderlineUnderlineWavyHeavy = '\002\001\000\033',
	WordWdUnderlineUnderlineDottedHeavy = '\002\001\000\024',
	WordWdUnderlineUnderlineDashHeavy = '\002\001\000\027',
	WordWdUnderlineUnderlineDotDashHeavy = '\002\001\000\031',
	WordWdUnderlineUnderlineDotDotDashHeavy = '\002\001\000\032',
	WordWdUnderlineUnderlineDashLong = '\002\001\000\'',
	WordWdUnderlineUnderlineDashLongHeavy = '\002\001\0007',
	WordWdUnderlineUnderlineWavyDouble = '\002\001\000+'
};
typedef enum WordWdUnderline WordWdUnderline;

enum WordWdEmphasisMark {
	WordWdEmphasisMarkEmphasisMarkNone = '\002\002\000\000',
	WordWdEmphasisMarkEmphasisMarkOverSolidCircle = '\002\002\000\001',
	WordWdEmphasisMarkEmphasisMarkOverComma = '\002\002\000\002',
	WordWdEmphasisMarkEmphasisMarkOverWhiteCircle = '\002\002\000\003',
	WordWdEmphasisMarkEmphasisMarkUnderSolidCircle = '\002\002\000\004'
};
typedef enum WordWdEmphasisMark WordWdEmphasisMark;

enum WordWdInternationalIndex {
	WordWdInternationalIndexListSeparator = '\002\003\000\021',
	WordWdInternationalIndexDecimalSeparator = '\002\003\000\022',
	WordWdInternationalIndexThousandsSeparator = '\002\003\000\023',
	WordWdInternationalIndexCurrencyCode = '\002\003\000\024',
	WordWdInternationalIndexTwentyFourHourClock = '\002\003\000\025',
	WordWdInternationalIndexInternationalAm = '\002\003\000\026',
	WordWdInternationalIndexInternationalPm = '\002\003\000\027',
	WordWdInternationalIndexTimeSeparator = '\002\003\000\030',
	WordWdInternationalIndexDateSeparator = '\002\003\000\031',
	WordWdInternationalIndexProductLanguageID = '\002\003\000\032'
};
typedef enum WordWdInternationalIndex WordWdInternationalIndex;

enum WordWdAutoMacros {
	WordWdAutoMacrosAutoExec = '\002\004\000\000',
	WordWdAutoMacrosAutoNew = '\002\004\000\001',
	WordWdAutoMacrosAutoOpen = '\002\004\000\002',
	WordWdAutoMacrosAutoClose = '\002\004\000\003',
	WordWdAutoMacrosAutoExit = '\002\004\000\004',
	WordWdAutoMacrosAutoSync = '\002\004\000\005'
};
typedef enum WordWdAutoMacros WordWdAutoMacros;

enum WordWdCaptionPosition {
	WordWdCaptionPositionCaptionPositionAbove = '\002\005\000\000',
	WordWdCaptionPositionCaptionPositionBelow = '\002\005\000\001'
};
typedef enum WordWdCaptionPosition WordWdCaptionPosition;

enum WordWdCountry {
	WordWdCountryUs = '\002\006\000\001',
	WordWdCountryCanada = '\002\006\000\002',
	WordWdCountryLatinAmerica = '\002\006\000\003',
	WordWdCountryNetherlands = '\002\006\000\037',
	WordWdCountryFrance = '\002\006\000!',
	WordWdCountrySpain = '\002\006\000\"',
	WordWdCountryItaly = '\002\006\000\'',
	WordWdCountryUk = '\002\006\000,',
	WordWdCountryDenmark = '\002\006\000-',
	WordWdCountrySweden = '\002\006\000.',
	WordWdCountryNorway = '\002\006\000/',
	WordWdCountryGermany = '\002\006\0001',
	WordWdCountryPeru = '\002\006\0003',
	WordWdCountryMexico = '\002\006\0004',
	WordWdCountryArgentina = '\002\006\0006',
	WordWdCountryBrazil = '\002\006\0007',
	WordWdCountryChile = '\002\006\0008',
	WordWdCountryVenezuela = '\002\006\000:',
	WordWdCountryJapan = '\002\006\000Q',
	WordWdCountryTaiwan = '\002\006\003v',
	WordWdCountryChina = '\002\006\000V',
	WordWdCountryKorea = '\002\006\000R',
	WordWdCountryFinland = '\002\006\001f',
	WordWdCountryIceland = '\002\006\001b'
};
typedef enum WordWdCountry WordWdCountry;

enum WordWdHeadingSeparator {
	WordWdHeadingSeparatorHeadingSeparatorNone = '\002\007\000\000',
	WordWdHeadingSeparatorHeadingSeparatorBlankLine = '\002\007\000\001',
	WordWdHeadingSeparatorHeadingSeparatorLetter = '\002\007\000\002',
	WordWdHeadingSeparatorHeadingSeparatorLetterLow = '\002\007\000\003',
	WordWdHeadingSeparatorHeadingSeparatorLetterFull = '\002\007\000\004'
};
typedef enum WordWdHeadingSeparator WordWdHeadingSeparator;

enum WordWdSeparatorType {
	WordWdSeparatorTypeSeparatorHyphen = '\002\010\000\000',
	WordWdSeparatorTypeSeparatorPeriod = '\002\010\000\001',
	WordWdSeparatorTypeSeparatorColon = '\002\010\000\002',
	WordWdSeparatorTypeSeparatorEmDash = '\002\010\000\003',
	WordWdSeparatorTypeSeparatorEnDash = '\002\010\000\004'
};
typedef enum WordWdSeparatorType WordWdSeparatorType;

enum WordWdPageNumberAlignment {
	WordWdPageNumberAlignmentAlignPageNumberLeft = '\002\011\000\000',
	WordWdPageNumberAlignmentAlignPageNumberCenter = '\002\011\000\001',
	WordWdPageNumberAlignmentAlignPageNumberRight = '\002\011\000\002',
	WordWdPageNumberAlignmentAlignPageNumberInside = '\002\011\000\003',
	WordWdPageNumberAlignmentAlignPageNumberOutside = '\002\011\000\004'
};
typedef enum WordWdPageNumberAlignment WordWdPageNumberAlignment;

enum WordWdBorderType {
	WordWdBorderTypeBorderTop = '\002\011\377\377',
	WordWdBorderTypeBorderLeft = '\002\011\377\376',
	WordWdBorderTypeBorderBottom = '\002\011\377\375',
	WordWdBorderTypeBorderRight = '\002\011\377\374',
	WordWdBorderTypeBorderHorizontal = '\002\011\377\373',
	WordWdBorderTypeBorderVertical = '\002\011\377\372',
	WordWdBorderTypeBorderDiagonalDown = '\002\011\377\371',
	WordWdBorderTypeBorderDiagonalUp = '\002\011\377\370'
};
typedef enum WordWdBorderType WordWdBorderType;

enum WordWdFramePosition {
	WordWdFramePositionFrameTop = '\001\373\275\301',
	WordWdFramePositionFrameLeft = '\001\373\275\302',
	WordWdFramePositionFrameBottom = '\001\373\275\303',
	WordWdFramePositionFrameRight = '\001\373\275\304',
	WordWdFramePositionFrameCenter = '\001\373\275\305',
	WordWdFramePositionFrameInside = '\001\373\275\306',
	WordWdFramePositionFrameOutside = '\001\373\275\307'
};
typedef enum WordWdFramePosition WordWdFramePosition;

enum WordWdAnimation {
	WordWdAnimationAnimationNone = '\002\014\000\000',
	WordWdAnimationAnimationLasVegasLights = '\002\014\000\001',
	WordWdAnimationAnimationBlinkingBackground = '\002\014\000\002',
	WordWdAnimationAnimationSparkleText = '\002\014\000\003',
	WordWdAnimationAnimationMarchingBlackAnts = '\002\014\000\004',
	WordWdAnimationAnimationMarchingRedAnts = '\002\014\000\005',
	WordWdAnimationAnimationShimmer = '\002\014\000\006'
};
typedef enum WordWdAnimation WordWdAnimation;

enum WordWdCharacterCase {
	WordWdCharacterCaseNextCase = '\002\014\377\377',
	WordWdCharacterCaseLowerCase = '\002\015\000\000',
	WordWdCharacterCaseUpperCase = '\002\015\000\001',
	WordWdCharacterCaseTitleWord = '\002\015\000\002',
	WordWdCharacterCaseTitleSentence = '\002\015\000\004',
	WordWdCharacterCaseToggleCase = '\002\015\000\005',
	WordWdCharacterCaseCaseHalfWidth = '\002\015\000\006',
	WordWdCharacterCaseCaseFullWidth = '\002\015\000\007',
	WordWdCharacterCaseCaseKatakana = '\002\015\000\010',
	WordWdCharacterCaseCaseHiragana = '\002\015\000\011'
};
typedef enum WordWdCharacterCase WordWdCharacterCase;

enum WordWdSummaryLength {
	WordWdSummaryLength10Sentences = '\002\016\377\376',
	WordWdSummaryLength20Sentences = '\002\016\377\375',
	WordWdSummaryLength100Words = '\002\016\377\374',
	WordWdSummaryLength500Words = '\002\016\377\373',
	WordWdSummaryLength10Percent = '\002\016\377\372',
	WordWdSummaryLength25Percent = '\002\016\377\371',
	WordWdSummaryLength50Percent = '\002\016\377\370',
	WordWdSummaryLength75Percent = '\002\016\377\367'
};
typedef enum WordWdSummaryLength WordWdSummaryLength;

enum WordWdStyleType {
	WordWdStyleTypeStyleTypeParagraph = '\002\020\000\001',
	WordWdStyleTypeStyleTypeCharacter = '\002\020\000\002',
	WordWdStyleTypeStyleTypeTable = '\002\020\000\003',
	WordWdStyleTypeStyleTypeList = '\002\020\000\004',
	WordWdStyleTypeStyleTypeParagraphOnly = '\002\020\000\005',
	WordWdStyleTypeStyleTypeLinked = '\002\020\000\006'
};
typedef enum WordWdStyleType WordWdStyleType;

enum WordWdUnits {
	WordWdUnitsACharacterItem = '\002\021\000\001',
	WordWdUnitsAWordItem = '\002\021\000\002',
	WordWdUnitsASentenceItem = '\002\021\000\003',
	WordWdUnitsAParagraphItem = '\002\021\000\004',
	WordWdUnitsALineItem = '\002\021\000\005',
	WordWdUnitsAStoryItem = '\002\021\000\006',
	WordWdUnitsAScreen = '\002\021\000\007',
	WordWdUnitsASection = '\002\021\000\010',
	WordWdUnitsAColumn = '\002\021\000\011',
	WordWdUnitsARow = '\002\021\000\012',
	WordWdUnitsAWindow = '\002\021\000\013',
	WordWdUnitsACell = '\002\021\000\014',
	WordWdUnitsACharacterFormatting = '\002\021\000\015',
	WordWdUnitsAParagraphFormatting = '\002\021\000\016',
	WordWdUnitsATable = '\002\021\000\017',
	WordWdUnitsAnItemUnit = '\002\021\000\020'
};
typedef enum WordWdUnits WordWdUnits;

enum WordWdGoToItem {
	WordWdGoToItemGotoABookmarkItem = '\002\021\377\377',
	WordWdGoToItemGotoASectionItem = '\002\022\000\000',
	WordWdGoToItemGotoAPageItem = '\002\022\000\001',
	WordWdGoToItemGotoATableItem = '\002\022\000\002',
	WordWdGoToItemGotoALineItem = '\002\022\000\003',
	WordWdGoToItemGotoAFootnoteItem = '\002\022\000\004',
	WordWdGoToItemGotoAnEndnoteItem = '\002\022\000\005',
	WordWdGoToItemGotoACommentItem = '\002\022\000\006',
	WordWdGoToItemGotoAFieldItem = '\002\022\000\007',
	WordWdGoToItemGotoAGraphic = '\002\022\000\010',
	WordWdGoToItemGotoAnObjectItem = '\002\022\000\011',
	WordWdGoToItemGotoAnEquation = '\002\022\000\012',
	WordWdGoToItemGotoAHeadingItem = '\002\022\000\013',
	WordWdGoToItemGotoAPercent = '\002\022\000\014',
	WordWdGoToItemGotoASpellingError = '\002\022\000\015',
	WordWdGoToItemGotoAGrammaticalError = '\002\022\000\016',
	WordWdGoToItemGotoAProofreadingError = '\002\022\000\017'
};
typedef enum WordWdGoToItem WordWdGoToItem;

enum WordWdGoToDirection {
	WordWdGoToDirectionTheFirstItem = '\002\023\000\001',
	WordWdGoToDirectionTheLastItem = '\002\022\377\377',
	WordWdGoToDirectionTheNextItem = '\002\023\000\002',
	WordWdGoToDirectionRelative = '\002\023\000\002',
	WordWdGoToDirectionThePreviousItem = '\002\023\000\003',
	WordWdGoToDirectionAbsolute = '\002\023\000\001'
};
typedef enum WordWdGoToDirection WordWdGoToDirection;

enum WordWdCollapseDirection {
	WordWdCollapseDirectionCollapseStart = '\002\024\000\001',
	WordWdCollapseDirectionCollapseEnd = '\002\024\000\000'
};
typedef enum WordWdCollapseDirection WordWdCollapseDirection;

enum WordWdRowHeightRule {
	WordWdRowHeightRuleRowHeightAuto = '\002\025\000\000',
	WordWdRowHeightRuleRowHeightAtLeast = '\002\025\000\001',
	WordWdRowHeightRuleRowHeightExactly = '\002\025\000\002'
};
typedef enum WordWdRowHeightRule WordWdRowHeightRule;

enum WordWdFrameSizeRule {
	WordWdFrameSizeRuleFrameAuto = '\002\026\000\000',
	WordWdFrameSizeRuleFrameAtLeast = '\002\026\000\001',
	WordWdFrameSizeRuleFrameExact = '\002\026\000\002'
};
typedef enum WordWdFrameSizeRule WordWdFrameSizeRule;

enum WordWdInsertCells {
	WordWdInsertCellsInsertCellsShiftRight = '\002\027\000\000',
	WordWdInsertCellsInsertCellsShiftDown = '\002\027\000\001',
	WordWdInsertCellsInsertCellsEntireRow = '\002\027\000\002',
	WordWdInsertCellsInsertCellsEntireColumn = '\002\027\000\003'
};
typedef enum WordWdInsertCells WordWdInsertCells;

enum WordWdDeleteCells {
	WordWdDeleteCellsDeleteCellsShiftLeft = '\002\030\000\000',
	WordWdDeleteCellsDeleteCellsShiftUp = '\002\030\000\001',
	WordWdDeleteCellsDeleteCellsEntireRow = '\002\030\000\002',
	WordWdDeleteCellsDeleteCellsEntireColumn = '\002\030\000\003'
};
typedef enum WordWdDeleteCells WordWdDeleteCells;

enum WordWdListApplyTo {
	WordWdListApplyToListApplyToWholeList = '\002\031\000\000',
	WordWdListApplyToListApplyToThisPointForward = '\002\031\000\001',
	WordWdListApplyToListApplyToSelection = '\002\031\000\002'
};
typedef enum WordWdListApplyTo WordWdListApplyTo;

enum WordWdAlertLevel {
	WordWdAlertLevelAlertsNone = '\002\032\000\000',
	WordWdAlertLevelAlertsMessageBox = '\002\031\377\376',
	WordWdAlertLevelAlertsAll = '\002\031\377\377'
};
typedef enum WordWdAlertLevel WordWdAlertLevel;

enum WordWdCursorType {
	WordWdCursorTypeCursorWait = '\002\033\000\000',
	WordWdCursorTypeCursorIbeam = '\002\033\000\001',
	WordWdCursorTypeCursorNormal = '\002\033\000\002',
	WordWdCursorTypeCursorNorthwestArrow = '\002\033\000\003'
};
typedef enum WordWdCursorType WordWdCursorType;

enum WordWdRulerStyle {
	WordWdRulerStyleAdjustNone = '\002\035\000\000',
	WordWdRulerStyleAdjustProportional = '\002\035\000\001',
	WordWdRulerStyleAdjustFirstColumn = '\002\035\000\002',
	WordWdRulerStyleAdjustSameWidth = '\002\035\000\003'
};
typedef enum WordWdRulerStyle WordWdRulerStyle;

enum WordWdParagraphAlignment {
	WordWdParagraphAlignmentAlignParagraphLeft = '\002\036\000\000',
	WordWdParagraphAlignmentAlignParagraphCenter = '\002\036\000\001',
	WordWdParagraphAlignmentAlignParagraphRight = '\002\036\000\002',
	WordWdParagraphAlignmentAlignParagraphJustify = '\002\036\000\003',
	WordWdParagraphAlignmentAlignParagraphDistribute = '\002\036\000\004',
	WordWdParagraphAlignmentAlignParagraphJustifyMed = '\002\036\000\005',
	WordWdParagraphAlignmentAlignParagraphJustifyHi = '\002\036\000\007',
	WordWdParagraphAlignmentAlignParagraphJustifyLow = '\002\036\000\010',
	WordWdParagraphAlignmentAlignParagraphThaiJustify = '\002\036\000\011'
};
typedef enum WordWdParagraphAlignment WordWdParagraphAlignment;

enum WordWdListLevelAlignment {
	WordWdListLevelAlignmentListLevelAlignLeft = '\002\037\000\000',
	WordWdListLevelAlignmentListLevelAlignCenter = '\002\037\000\001',
	WordWdListLevelAlignmentListLevelAlignRight = '\002\037\000\002'
};
typedef enum WordWdListLevelAlignment WordWdListLevelAlignment;

enum WordWdRowAlignment {
	WordWdRowAlignmentAlignRowLeft = '\002 \000\000',
	WordWdRowAlignmentAlignRowCenter = '\002 \000\001',
	WordWdRowAlignmentAlignRowRight = '\002 \000\002'
};
typedef enum WordWdRowAlignment WordWdRowAlignment;

enum WordWdTabAlignment {
	WordWdTabAlignmentAlignTabLeft = '\002!\000\000',
	WordWdTabAlignmentAlignTabCenter = '\002!\000\001',
	WordWdTabAlignmentAlignTabRight = '\002!\000\002',
	WordWdTabAlignmentAlignTabDecimal = '\002!\000\003',
	WordWdTabAlignmentAlignTabBar = '\002!\000\004',
	WordWdTabAlignmentAlignTabList = '\002!\000\006'
};
typedef enum WordWdTabAlignment WordWdTabAlignment;

enum WordWdVerticalAlignment {
	WordWdVerticalAlignmentAlignVerticalTop = '\002\"\000\000',
	WordWdVerticalAlignmentAlignVerticalCenter = '\002\"\000\001',
	WordWdVerticalAlignmentAlignVerticalJustify = '\002\"\000\002',
	WordWdVerticalAlignmentAlignVerticalBottom = '\002\"\000\003'
};
typedef enum WordWdVerticalAlignment WordWdVerticalAlignment;

enum WordWdCellVerticalAlignment {
	WordWdCellVerticalAlignmentCellAlignVerticalTop = '\002#\000\000',
	WordWdCellVerticalAlignmentCellAlignVerticalCenter = '\002#\000\001',
	WordWdCellVerticalAlignmentCellAlignVerticalBottom = '\002#\000\003'
};
typedef enum WordWdCellVerticalAlignment WordWdCellVerticalAlignment;

enum WordWdRangeAnnotationAlignment {
	WordWdRangeAnnotationAlignmentRangeAnnotationAlignmentCenter = '\002$\000\000',
	WordWdRangeAnnotationAlignmentZeroOneZero = '\002$\000\001',
	WordWdRangeAnnotationAlignmentOneTwoOne = '\002$\000\002',
	WordWdRangeAnnotationAlignmentRangeAnnotationAlignmentLeft = '\002$\000\003',
	WordWdRangeAnnotationAlignmentRangeAnnotationAlignmentRight = '\002$\000\004'
};
typedef enum WordWdRangeAnnotationAlignment WordWdRangeAnnotationAlignment;

enum WordWdTrailingCharacter {
	WordWdTrailingCharacterTrailingTab = '\002%\000\000',
	WordWdTrailingCharacterTrailingSpace = '\002%\000\001',
	WordWdTrailingCharacterTrailingNone = '\002%\000\002'
};
typedef enum WordWdTrailingCharacter WordWdTrailingCharacter;

enum WordWdListGalleryType {
	WordWdListGalleryTypeBulletGallery = '\002&\000\001',
	WordWdListGalleryTypeNumberGallery = '\002&\000\002',
	WordWdListGalleryTypeOutlineNumberGallery = '\002&\000\003'
};
typedef enum WordWdListGalleryType WordWdListGalleryType;

enum WordWdListNumberStyle {
	WordWdListNumberStyleListNumberStyleArabic = '\002\'\000\000',
	WordWdListNumberStyleListNumberStyleUppercaseRoman = '\002\'\000\001',
	WordWdListNumberStyleListNumberStyleLowercaseRoman = '\002\'\000\002',
	WordWdListNumberStyleListNumberStyleUppercaseLetter = '\002\'\000\003',
	WordWdListNumberStyleListNumberStyleLowercaseLetter = '\002\'\000\004',
	WordWdListNumberStyleListNumberStyleOrdinal = '\002\'\000\005',
	WordWdListNumberStyleListNumberStyleCardinalText = '\002\'\000\006',
	WordWdListNumberStyleListNumberStyleOrdinalText = '\002\'\000\007',
	WordWdListNumberStyleListNumberStyleKanji = '\002\'\000\012',
	WordWdListNumberStyleListNumberStyleKanjiDigit = '\002\'\000\013',
	WordWdListNumberStyleListNumberStyleAiueoHalfWidth = '\002\'\000\014',
	WordWdListNumberStyleListNumberStyleIrohaHalfWidth = '\002\'\000\015',
	WordWdListNumberStyleListNumberStyleArabicFullWidth = '\002\'\000\016',
	WordWdListNumberStyleListNumberStyleKanjiTraditional = '\002\'\000\020',
	WordWdListNumberStyleListNumberStyleKanjiTraditional2 = '\002\'\000\021',
	WordWdListNumberStyleListNumberStyleNumberInCircle = '\002\'\000\022',
	WordWdListNumberStyleListNumberStyleAiueo = '\002\'\000\024',
	WordWdListNumberStyleListNumberStyleIroha = '\002\'\000\025',
	WordWdListNumberStyleListNumberStyleArabicLz = '\002\'\000\026',
	WordWdListNumberStyleListNumberStyleBullet = '\002\'\000\027',
	WordWdListNumberStyleListNumberStyleGanada = '\002\'\000\030',
	WordWdListNumberStyleListNumberStyleChosung = '\002\'\000\031',
	WordWdListNumberStyleListNumberStyleGbnum1 = '\002\'\000\032',
	WordWdListNumberStyleListNumberStyleGbnum2 = '\002\'\000\033',
	WordWdListNumberStyleListNumberStyleGbnum3 = '\002\'\000\034',
	WordWdListNumberStyleListNumberStyleGbnum4 = '\002\'\000\035',
	WordWdListNumberStyleListNumberStyleZodiac1 = '\002\'\000\036',
	WordWdListNumberStyleListNumberStyleZodiac2 = '\002\'\000\037',
	WordWdListNumberStyleListNumberStyleZodiac3 = '\002\'\000 ',
	WordWdListNumberStyleListNumberStyleTradChinNum1 = '\002\'\000!',
	WordWdListNumberStyleListNumberStyleTradChinNum2 = '\002\'\000\"',
	WordWdListNumberStyleListNumberStyleTradChinNum3 = '\002\'\000#',
	WordWdListNumberStyleListNumberStyleTradChinNum4 = '\002\'\000$',
	WordWdListNumberStyleListNumberStyleSimpChinNum1 = '\002\'\000%',
	WordWdListNumberStyleListNumberStyleSimpChinNum2 = '\002\'\000&',
	WordWdListNumberStyleListNumberStyleSimpChinNum3 = '\002\'\000\'',
	WordWdListNumberStyleListNumberStyleSimpChinNum4 = '\002\'\000(',
	WordWdListNumberStyleListNumberStyleHanjaRead = '\002\'\000)',
	WordWdListNumberStyleListNumberStyleHanjaReadDigit = '\002\'\000*',
	WordWdListNumberStyleListNumberStyleHangul = '\002\'\000+',
	WordWdListNumberStyleListNumberStyleHanja = '\002\'\000,',
	WordWdListNumberStyleListNumberStyleHebrew1 = '\002\'\000-',
	WordWdListNumberStyleListNumberStyleArabic1 = '\002\'\000.',
	WordWdListNumberStyleListNumberStyleHebrew2 = '\002\'\000/',
	WordWdListNumberStyleListNumberStyleArabic2 = '\002\'\0000',
	WordWdListNumberStyleListNumberStyleHindiLetter1 = '\002\'\0001',
	WordWdListNumberStyleListNumberStyleHindiLetter2 = '\002\'\0002',
	WordWdListNumberStyleListNumberStyleHindiArabic = '\002\'\0003',
	WordWdListNumberStyleListNumberStyleHindiCardinalText = '\002\'\0004',
	WordWdListNumberStyleListNumberStyleThaiLetter = '\002\'\0005',
	WordWdListNumberStyleListNumberStyleThaiArabic = '\002\'\0006',
	WordWdListNumberStyleListNumberStyleThaiCardinalText = '\002\'\0007',
	WordWdListNumberStyleListNumberStyleVietCardinalText = '\002\'\0008',
	WordWdListNumberStyleListNumberStyleLowercaseRussian = '\002\'\000:',
	WordWdListNumberStyleListNumberStyleUppercaseRussian = '\002\'\000;',
	WordWdListNumberStyleListNumberStyleLowercaseGreek = '\002\'\000<',
	WordWdListNumberStyleListNumberStyleUppercaseGreek = '\002\'\000=',
	WordWdListNumberStyleListNumberStyleArabicLz2 = '\002\'\000>',
	WordWdListNumberStyleListNumberStyleArabicLz3 = '\002\'\000\?',
	WordWdListNumberStyleListNumberStyleArabicLz4 = '\002\'\000@',
	WordWdListNumberStyleListNumberStyleLowercaseTurkish = '\002\'\000A',
	WordWdListNumberStyleListNumberStyleUppercaseTurkish = '\002\'\000B',
	WordWdListNumberStyleListNumberStyleLowercaseBulgarian = '\002\'\000C',
	WordWdListNumberStyleListNumberStyleUppercaseBulgarian = '\002\'\000D',
	WordWdListNumberStyleListNumberStylePictureBullet = '\002\'\000\371',
	WordWdListNumberStyleListNumberStyleLegal = '\002\'\000\375',
	WordWdListNumberStyleListNumberStyleLegalLz = '\002\'\000\376',
	WordWdListNumberStyleListNumberStyleNone = '\002\'\000\377'
};
typedef enum WordWdListNumberStyle WordWdListNumberStyle;

enum WordWdNoteNumberStyle {
	WordWdNoteNumberStyleNoteNumberStyleArabic = '\002(\000\000',
	WordWdNoteNumberStyleNoteNumberStyleUppercaseRoman = '\002(\000\001',
	WordWdNoteNumberStyleNoteNumberStyleLowercaseRoman = '\002(\000\002',
	WordWdNoteNumberStyleNoteNumberStyleUppercaseLetter = '\002(\000\003',
	WordWdNoteNumberStyleNoteNumberStyleLowercaseLetter = '\002(\000\004',
	WordWdNoteNumberStyleNoteNumberStyleSymbol = '\002(\000\011',
	WordWdNoteNumberStyleNoteNumberStyleArabicFullWidth = '\002(\000\016',
	WordWdNoteNumberStyleNoteNumberStyleKanji = '\002(\000\012',
	WordWdNoteNumberStyleNoteNumberStyleKanjiDigit = '\002(\000\013',
	WordWdNoteNumberStyleNoteNumberStyleKanjiTraditional = '\002(\000\020',
	WordWdNoteNumberStyleNoteNumberStyleNumberInCircle = '\002(\000\022',
	WordWdNoteNumberStyleNoteNumberStyleHanjaRead = '\002(\000)',
	WordWdNoteNumberStyleNoteNumberStyleHanjaReadDigit = '\002(\000*',
	WordWdNoteNumberStyleNoteNumberStyleTradChinNum1 = '\002(\000!',
	WordWdNoteNumberStyleNoteNumberStyleTradChinNum2 = '\002(\000\"',
	WordWdNoteNumberStyleNoteNumberStyleSimpChinNum1 = '\002(\000%',
	WordWdNoteNumberStyleNoteNumberStyleSimpChinNum2 = '\002(\000&',
	WordWdNoteNumberStyleNoteNumberStyleHebrewLetter1 = '\002(\000-',
	WordWdNoteNumberStyleNoteNumberStyleArabicLetter1 = '\002(\000.',
	WordWdNoteNumberStyleNoteNumberStyleHebrewLetter2 = '\002(\000/',
	WordWdNoteNumberStyleNoteNumberStyleArabicLetter2 = '\002(\0000',
	WordWdNoteNumberStyleNoteNumberStyleHindiLetter1 = '\002(\0001',
	WordWdNoteNumberStyleNoteNumberStyleHindiLetter2 = '\002(\0002',
	WordWdNoteNumberStyleNoteNumberStyleHindiArabic = '\002(\0003',
	WordWdNoteNumberStyleNoteNumberStyleHindiCardinalText = '\002(\0004',
	WordWdNoteNumberStyleNoteNumberStyleThaiLetter = '\002(\0005',
	WordWdNoteNumberStyleNoteNumberStyleThaiArabic = '\002(\0006',
	WordWdNoteNumberStyleNoteNumberStyleThaiCardinalText = '\002(\0007',
	WordWdNoteNumberStyleNoteNumberStyleVietCardinalText = '\002(\0008'
};
typedef enum WordWdNoteNumberStyle WordWdNoteNumberStyle;

enum WordWdCaptionNumberStyle {
	WordWdCaptionNumberStyleCaptionNumberStyleArabic = '\002)\000\000',
	WordWdCaptionNumberStyleCaptionNumberStyleUppercaseRoman = '\002)\000\001',
	WordWdCaptionNumberStyleCaptionNumberStyleLowercaseRoman = '\002)\000\002',
	WordWdCaptionNumberStyleCaptionNumberStyleUppercaseLetter = '\002)\000\003',
	WordWdCaptionNumberStyleCaptionNumberStyleLowercaseLetter = '\002)\000\004',
	WordWdCaptionNumberStyleCaptionNumberStyleArabicFullWidth = '\002)\000\016',
	WordWdCaptionNumberStyleCaptionNumberStyleKanji = '\002)\000\012',
	WordWdCaptionNumberStyleCaptionNumberStyleKanjiDigit = '\002)\000\013',
	WordWdCaptionNumberStyleCaptionNumberStyleKanjiTraditional = '\002)\000\020',
	WordWdCaptionNumberStyleCaptionNumberStyleNumberInCircle = '\002)\000\022',
	WordWdCaptionNumberStyleCaptionNumberStyleGanada = '\002)\000\030',
	WordWdCaptionNumberStyleCaptionNumberStyleChosung = '\002)\000\031',
	WordWdCaptionNumberStyleCaptionNumberStyleZodiac1 = '\002)\000\036',
	WordWdCaptionNumberStyleCaptionNumberStyleZodiac2 = '\002)\000\037',
	WordWdCaptionNumberStyleCaptionNumberStyleHanjaRead = '\002)\000)',
	WordWdCaptionNumberStyleCaptionNumberStyleHanjaReadDigit = '\002)\000*',
	WordWdCaptionNumberStyleCaptionNumberStyleTradChinNum2 = '\002)\000\"',
	WordWdCaptionNumberStyleCaptionNumberStyleTradChinNum3 = '\002)\000#',
	WordWdCaptionNumberStyleCaptionNumberStyleSimpChinNum2 = '\002)\000&',
	WordWdCaptionNumberStyleCaptionNumberStyleSimpChinNum3 = '\002)\000\'',
	WordWdCaptionNumberStyleCaptionNumberStyleHebrewLetter1 = '\002)\000-',
	WordWdCaptionNumberStyleCaptionNumberStyleArabicLetter1 = '\002)\000.',
	WordWdCaptionNumberStyleCaptionNumberStyleHebrewLetter2 = '\002)\000/',
	WordWdCaptionNumberStyleCaptionNumberStyleArabicLetter2 = '\002)\0000',
	WordWdCaptionNumberStyleCaptionNumberStyleHindiLetter1 = '\002)\0001',
	WordWdCaptionNumberStyleCaptionNumberStyleHindiLetter2 = '\002)\0002',
	WordWdCaptionNumberStyleCaptionNumberStyleHindiArabic = '\002)\0003',
	WordWdCaptionNumberStyleCaptionNumberStyleHindiCardinalText = '\002)\0004',
	WordWdCaptionNumberStyleCaptionNumberStyleThaiLetter = '\002)\0005',
	WordWdCaptionNumberStyleCaptionNumberStyleThaiArabic = '\002)\0006',
	WordWdCaptionNumberStyleCaptionNumberStyleThaiCardinalText = '\002)\0007',
	WordWdCaptionNumberStyleCaptionNumberStyleVietCardinalText = '\002)\0008'
};
typedef enum WordWdCaptionNumberStyle WordWdCaptionNumberStyle;

enum WordWdPageNumberStyle {
	WordWdPageNumberStylePageNumberStyleArabic = '\002*\000\000',
	WordWdPageNumberStylePageNumberStyleUppercaseRoman = '\002*\000\001',
	WordWdPageNumberStylePageNumberStyleLowercaseRoman = '\002*\000\002',
	WordWdPageNumberStylePageNumberStyleUppercaseLetter = '\002*\000\003',
	WordWdPageNumberStylePageNumberStyleLowercaseLetter = '\002*\000\004',
	WordWdPageNumberStylePageNumberStyleArabicFullWidth = '\002*\000\016',
	WordWdPageNumberStylePageNumberStyleKanji = '\002*\000\012',
	WordWdPageNumberStylePageNumberStyleKanjiDigit = '\002*\000\013',
	WordWdPageNumberStylePageNumberStyleKanjiTraditional = '\002*\000\020',
	WordWdPageNumberStylePageNumberStyleNumberInCircle = '\002*\000\022',
	WordWdPageNumberStylePageNumberStyleHanjaRead = '\002*\000)',
	WordWdPageNumberStylePageNumberStyleHanjaReadDigit = '\002*\000*',
	WordWdPageNumberStylePageNumberStyleTradChinNum1 = '\002*\000!',
	WordWdPageNumberStylePageNumberStyleTradChinNum2 = '\002*\000\"',
	WordWdPageNumberStylePageNumberStyleSimpChinNum1 = '\002*\000%',
	WordWdPageNumberStylePageNumberStyleSimpChinNum2 = '\002*\000&',
	WordWdPageNumberStylePageNumberStyleHebrewLetter1 = '\002*\000-',
	WordWdPageNumberStylePageNumberStyleArabicLetter1 = '\002*\000.',
	WordWdPageNumberStylePageNumberStyleHebrewLetter2 = '\002*\000/',
	WordWdPageNumberStylePageNumberStyleArabicLetter2 = '\002*\0000',
	WordWdPageNumberStylePageNumberStyleHindiLetter1 = '\002*\0001',
	WordWdPageNumberStylePageNumberStyleHindiLetter2 = '\002*\0002',
	WordWdPageNumberStylePageNumberStyleHindiArabic = '\002*\0003',
	WordWdPageNumberStylePageNumberStyleHindiCardinalText = '\002*\0004',
	WordWdPageNumberStylePageNumberStyleThaiLetter = '\002*\0005',
	WordWdPageNumberStylePageNumberStyleThaiArabic = '\002*\0006',
	WordWdPageNumberStylePageNumberStyleThaiCardinalText = '\002*\0007',
	WordWdPageNumberStylePageNumberStyleVietCardinalText = '\002*\0008',
	WordWdPageNumberStylePageNumberStyleNumberInDash = '\002*\0009'
};
typedef enum WordWdPageNumberStyle WordWdPageNumberStyle;

enum WordWdStatistic {
	WordWdStatisticStatisticWords = '\002+\000\000',
	WordWdStatisticStatisticLines = '\002+\000\001',
	WordWdStatisticStatisticPages = '\002+\000\002',
	WordWdStatisticStatisticCharacters = '\002+\000\003',
	WordWdStatisticStatisticParagraphs = '\002+\000\004',
	WordWdStatisticStatisticCharactersWithSpaces = '\002+\000\005',
	WordWdStatisticStatisticEastAsianCharacters = '\002+\000\006'
};
typedef enum WordWdStatistic WordWdStatistic;

enum WordWdBuiltInProperty {
	WordWdBuiltInPropertyPropertyTitle = '\002,\000\001',
	WordWdBuiltInPropertyPropertySubject = '\002,\000\002',
	WordWdBuiltInPropertyPropertyAuthor = '\002,\000\003',
	WordWdBuiltInPropertyPropertyKeywords = '\002,\000\004',
	WordWdBuiltInPropertyPropertyComments = '\002,\000\005',
	WordWdBuiltInPropertyPropertyTemplate = '\002,\000\006',
	WordWdBuiltInPropertyPropertyLastAuthor = '\002,\000\007',
	WordWdBuiltInPropertyPropertyRevision = '\002,\000\010',
	WordWdBuiltInPropertyPropertyAppName = '\002,\000\011',
	WordWdBuiltInPropertyPropertyTimeLastPrinted = '\002,\000\012',
	WordWdBuiltInPropertyPropertyTimeCreated = '\002,\000\013',
	WordWdBuiltInPropertyPropertyTimeLastSaved = '\002,\000\014',
	WordWdBuiltInPropertyPropertyVbatotalEdit = '\002,\000\015',
	WordWdBuiltInPropertyPropertyPages = '\002,\000\016',
	WordWdBuiltInPropertyPropertyWords = '\002,\000\017',
	WordWdBuiltInPropertyPropertyCharacters = '\002,\000\020',
	WordWdBuiltInPropertyPropertySecurity = '\002,\000\021',
	WordWdBuiltInPropertyPropertyCategory = '\002,\000\022',
	WordWdBuiltInPropertyPropertyFormat = '\002,\000\023',
	WordWdBuiltInPropertyPropertyManager = '\002,\000\024',
	WordWdBuiltInPropertyPropertyCompany = '\002,\000\025',
	WordWdBuiltInPropertyPropertyBytes = '\002,\000\026',
	WordWdBuiltInPropertyPropertyLines = '\002,\000\027',
	WordWdBuiltInPropertyPropertyParas = '\002,\000\030',
	WordWdBuiltInPropertyPropertySlides = '\002,\000\031',
	WordWdBuiltInPropertyPropertyNotes = '\002,\000\032',
	WordWdBuiltInPropertyPropertyHiddenSlides = '\002,\000\033',
	WordWdBuiltInPropertyPropertyMmclips = '\002,\000\034',
	WordWdBuiltInPropertyPropertyHyperlinkBase = '\002,\000\035',
	WordWdBuiltInPropertyPropertyCharsWspaces = '\002,\000\036'
};
typedef enum WordWdBuiltInProperty WordWdBuiltInProperty;

enum WordWdLineSpacing {
	WordWdLineSpacingLineSpaceSingle = '\002-\000\000',
	WordWdLineSpacingLineSpace1Pt5 = '\002-\000\001',
	WordWdLineSpacingLineSpaceDouble = '\002-\000\002',
	WordWdLineSpacingLineSpaceAtLeast = '\002-\000\003',
	WordWdLineSpacingLineSpaceExactly = '\002-\000\004',
	WordWdLineSpacingLineSpaceMultiple = '\002-\000\005'
};
typedef enum WordWdLineSpacing WordWdLineSpacing;

enum WordWdNumberType {
	WordWdNumberTypeNumberParagraph = '\002.\000\001',
	WordWdNumberTypeNumberListnum = '\002.\000\002',
	WordWdNumberTypeNumberAllNumbers = '\002.\000\003'
};
typedef enum WordWdNumberType WordWdNumberType;

enum WordWdListType {
	WordWdListTypeListNoNumbering = '\002/\000\000',
	WordWdListTypeListListnumOnly = '\002/\000\001',
	WordWdListTypeListBullet = '\002/\000\002',
	WordWdListTypeListSimpleNumbering = '\002/\000\003',
	WordWdListTypeListOutlineNumbering = '\002/\000\004',
	WordWdListTypeListMixedNumbering = '\002/\000\005',
	WordWdListTypeListPictureBullet = '\002/\000\006'
};
typedef enum WordWdListType WordWdListType;

enum WordWdStoryType {
	WordWdStoryTypeMainTextStory = '\0020\000\001',
	WordWdStoryTypeFootnotesStory = '\0020\000\002',
	WordWdStoryTypeEndnotesStory = '\0020\000\003',
	WordWdStoryTypeCommentsStory = '\0020\000\004',
	WordWdStoryTypeTextFrameStory = '\0020\000\005',
	WordWdStoryTypeEvenPagesHeaderStory = '\0020\000\006',
	WordWdStoryTypePrimaryHeaderStory = '\0020\000\007',
	WordWdStoryTypeEvenPagesFooterStory = '\0020\000\010',
	WordWdStoryTypePrimaryFooterStory = '\0020\000\011',
	WordWdStoryTypeFirstPageHeaderStory = '\0020\000\012',
	WordWdStoryTypeFirstPageFooterStory = '\0020\000\013',
	WordWdStoryTypeFootnoteSeparatorStory = '\0020\000\014',
	WordWdStoryTypeFootnoteContinuationSeparatorStory = '\0020\000\015',
	WordWdStoryTypeFootnoteContinuationNoticeStory = '\0020\000\016',
	WordWdStoryTypeEndnoteSeparatorStory = '\0020\000\017',
	WordWdStoryTypeEndnoteContinuationSeparatorStory = '\0020\000\020',
	WordWdStoryTypeEndnoteContinuationNoticeStory = '\0020\000\021'
};
typedef enum WordWdStoryType WordWdStoryType;

enum WordWdSaveFormat {
	WordWdSaveFormatFormatDocument97 = '\0021\000\000',
	WordWdSaveFormatFormatTemplate97 = '\0021\000\001',
	WordWdSaveFormatFormatText = '\0021\000\002',
	WordWdSaveFormatFormatTextLineBreaks = '\0021\000\003',
	WordWdSaveFormatFormatDostext = '\0021\000\004',
	WordWdSaveFormatFormatDostextLineBreaks = '\0021\000\005',
	WordWdSaveFormatFormatRtf = '\0021\000\006',
	WordWdSaveFormatFormatUnicodeText = '\0021\000\007',
	WordWdSaveFormatFormatHTML = '\0021\000\010',
	WordWdSaveFormatFormatWebArchive = '\0021\000\011',
	WordWdSaveFormatFormatFilteredHTML = '\0021\000\012',
	WordWdSaveFormatFormatXml = '\0021\000\013',
	WordWdSaveFormatFormatDocument = '\0021\000\014',
	WordWdSaveFormatFormatDocumentME = '\0021\000\015',
	WordWdSaveFormatFormatTemplate = '\0021\000\016',
	WordWdSaveFormatFormatTemplateME = '\0021\000\017',
	WordWdSaveFormatFormatDocumentDefault = '\0021\000\020',
	WordWdSaveFormatFormatPDF = '\0021\000\021',
	WordWdSaveFormatFormat = '\0021\000\022',
	WordWdSaveFormatFormatFlatDocument = '\0021\000\023',
	WordWdSaveFormatFormatFlatDocumentME = '\0021\000\024',
	WordWdSaveFormatFormatFlatTemplate = '\0021\000\025',
	WordWdSaveFormatFormatFlatTemplateME = '\0021\000\026',
	WordWdSaveFormatFormatODT = '\0021\000\027',
	WordWdSaveFormatFormatStrictOpenXmldocument = '\0021\000\030'
};
typedef enum WordWdSaveFormat WordWdSaveFormat;

enum WordWdOpenFormat {
	WordWdOpenFormatOpenFormatAuto = '\0022\000\000',
	WordWdOpenFormatOpenFormatDocument = '\0022\000\001',
	WordWdOpenFormatOpenFormatDocument97 = '\0022\000\001',
	WordWdOpenFormatOpenFormatTemplate = '\0022\000\002',
	WordWdOpenFormatOpenFormatTemplate97 = '\0022\000\002',
	WordWdOpenFormatOpenFormatRtf = '\0022\000\003',
	WordWdOpenFormatOpenFormatText = '\0022\000\004',
	WordWdOpenFormatOpenFormatUnicodeText = '\0022\000\005',
	WordWdOpenFormatOpenFormatEncodedText = '\0022\000\005',
	WordWdOpenFormatOpenFormatAllWord = '\0022\000\006',
	WordWdOpenFormatOpenFormatWebPages = '\0022\000\007',
	WordWdOpenFormatOpenFormatXml = '\0022\000\010',
	WordWdOpenFormatOpenFormatXmldocument = '\0022\000\011',
	WordWdOpenFormatOpenFormatXmldocumentMacroEnabled = '\0022\000\012',
	WordWdOpenFormatOpenFormatXmltemplate = '\0022\000\013',
	WordWdOpenFormatOpenFormatXmltemplateMacroEnabled = '\0022\000\014',
	WordWdOpenFormatOpenFormatAllWordTemplates = '\0022\000\015',
	WordWdOpenFormatOpenFormatXmldocumentSerialized = '\0022\000\016',
	WordWdOpenFormatOpenFormatXmldocumentMacroEnabledSerialized = '\0022\000\017',
	WordWdOpenFormatOpenFormatXmltemplateSerialized = '\0022\000\020',
	WordWdOpenFormatOpenFormatXmltemplateMacroEnabledSerialized = '\0022\000\021',
	WordWdOpenFormatOpenFormatOpenDocumentText = '\0022\000\022'
};
typedef enum WordWdOpenFormat WordWdOpenFormat;

enum WordWdHeaderFooterIndex {
	WordWdHeaderFooterIndexHeaderFooterPrimary = '\0023\000\001',
	WordWdHeaderFooterIndexHeaderFooterFirstPage = '\0023\000\002',
	WordWdHeaderFooterIndexHeaderFooterEvenPages = '\0023\000\003'
};
typedef enum WordWdHeaderFooterIndex WordWdHeaderFooterIndex;

enum WordWdTocFormat {
	WordWdTocFormatToctemplate = '\0024\000\000',
	WordWdTocFormatTocclassic = '\0024\000\001',
	WordWdTocFormatTocdistinctive = '\0024\000\002',
	WordWdTocFormatTocfancy = '\0024\000\003',
	WordWdTocFormatTocmodern = '\0024\000\004',
	WordWdTocFormatTocformal = '\0024\000\005',
	WordWdTocFormatTocsimple = '\0024\000\006'
};
typedef enum WordWdTocFormat WordWdTocFormat;

enum WordWdTofFormat {
	WordWdTofFormatToftemplate = '\0025\000\000',
	WordWdTofFormatTofclassic = '\0025\000\001',
	WordWdTofFormatTofdistinctive = '\0025\000\002',
	WordWdTofFormatTofcentered = '\0025\000\003',
	WordWdTofFormatTofformal = '\0025\000\004',
	WordWdTofFormatTofsimple = '\0025\000\005'
};
typedef enum WordWdTofFormat WordWdTofFormat;

enum WordWdToaFormat {
	WordWdToaFormatToatemplate = '\0026\000\000',
	WordWdToaFormatToaclassic = '\0026\000\001',
	WordWdToaFormatToadistinctive = '\0026\000\002',
	WordWdToaFormatToaformal = '\0026\000\003',
	WordWdToaFormatToasimple = '\0026\000\004'
};
typedef enum WordWdToaFormat WordWdToaFormat;

enum WordWdLineStyle {
	WordWdLineStyleLineStyleNone = '\0027\000\000',
	WordWdLineStyleLineStyleSingle = '\0027\000\001',
	WordWdLineStyleLineStyleDot = '\0027\000\002',
	WordWdLineStyleLineStyleDashSmallGap = '\0027\000\003',
	WordWdLineStyleLineStyleDashLargeGap = '\0027\000\004',
	WordWdLineStyleLineStyleDashDot = '\0027\000\005',
	WordWdLineStyleLineStyleDashDotDot = '\0027\000\006',
	WordWdLineStyleLineStyleDouble = '\0027\000\007',
	WordWdLineStyleLineStyleTriple = '\0027\000\010',
	WordWdLineStyleLineStyleThinThickSmallGap = '\0027\000\011',
	WordWdLineStyleLineStyleThickThinSmallGap = '\0027\000\012',
	WordWdLineStyleLineStyleThinThickThinSmallGap = '\0027\000\013',
	WordWdLineStyleLineStyleThinThickMedGap = '\0027\000\014',
	WordWdLineStyleLineStyleThickThinMedGap = '\0027\000\015',
	WordWdLineStyleLineStyleThinThickThinMedGap = '\0027\000\016',
	WordWdLineStyleLineStyleThinThickLargeGap = '\0027\000\017',
	WordWdLineStyleLineStyleThickThinLargeGap = '\0027\000\020',
	WordWdLineStyleLineStyleThinThickThinLargeGap = '\0027\000\021',
	WordWdLineStyleLineStyleSingleWavy = '\0027\000\022',
	WordWdLineStyleLineStyleDoubleWavy = '\0027\000\023',
	WordWdLineStyleLineStyleDashDotStroked = '\0027\000\024',
	WordWdLineStyleLineStyleEmboss_3D = '\0027\000\025',
	WordWdLineStyleLineStyleEngrave_3D = '\0027\000\026',
	WordWdLineStyleLineStyleOutset = '\0027\000\027',
	WordWdLineStyleLineStyleInset = '\0027\000\030'
};
typedef enum WordWdLineStyle WordWdLineStyle;

enum WordWdLineWidth {
	WordWdLineWidthLineWidth25Point = '\0028\000\002',
	WordWdLineWidthLineWidth50Point = '\0028\000\004',
	WordWdLineWidthLineWidth75Point = '\0028\000\006',
	WordWdLineWidthLineWidth100Point = '\0028\000\010',
	WordWdLineWidthLineWidth150Point = '\0028\000\014',
	WordWdLineWidthLineWidth225Point = '\0028\000\022',
	WordWdLineWidthLineWidth300Point = '\0028\000\030',
	WordWdLineWidthLineWidth450Point = '\0028\000$',
	WordWdLineWidthLineWidth600Point = '\0028\0000'
};
typedef enum WordWdLineWidth WordWdLineWidth;

enum WordWdBreakType {
	WordWdBreakTypeSectionBreakNextPage = '\0029\000\002',
	WordWdBreakTypeSectionBreakContinuous = '\0029\000\003',
	WordWdBreakTypeSectionBreakEvenPage = '\0029\000\004',
	WordWdBreakTypeSectionBreakOddPage = '\0029\000\005',
	WordWdBreakTypeLineBreak = '\0029\000\006',
	WordWdBreakTypePageBreak = '\0029\000\007',
	WordWdBreakTypeColumnBreak = '\0029\000\010',
	WordWdBreakTypeLineBreakClearLeft = '\0029\000\011',
	WordWdBreakTypeLineBreakClearRight = '\0029\000\012',
	WordWdBreakTypeTextWrappingBreak = '\0029\000\013'
};
typedef enum WordWdBreakType WordWdBreakType;

enum WordWdPageType {
	WordWdPageTypeContinuedMaster = '\002\311\000\000'
};
typedef enum WordWdPageType WordWdPageType;

enum WordWdTabLeader {
	WordWdTabLeaderTabLeaderSpaces = '\002:\000\000',
	WordWdTabLeaderTabLeaderDots = '\002:\000\001',
	WordWdTabLeaderTabLeaderDashes = '\002:\000\002',
	WordWdTabLeaderTabLeaderLines = '\002:\000\003',
	WordWdTabLeaderTabLeaderHeavy = '\002:\000\004',
	WordWdTabLeaderTabLeaderMiddleDot = '\002:\000\005'
};
typedef enum WordWdTabLeader WordWdTabLeader;

enum WordWdMeasurementUnits {
	WordWdMeasurementUnitsInches = '\002;\000\000',
	WordWdMeasurementUnitsCentimeters = '\002;\000\001',
	WordWdMeasurementUnitsMillimeters = '\002;\000\002',
	WordWdMeasurementUnitsPoints = '\002;\000\003',
	WordWdMeasurementUnitsPicas = '\002;\000\004'
};
typedef enum WordWdMeasurementUnits WordWdMeasurementUnits;

enum WordWdDropPosition {
	WordWdDropPositionDropNone = '\002<\000\000',
	WordWdDropPositionDropNormal = '\002<\000\001',
	WordWdDropPositionDropMargin = '\002<\000\002'
};
typedef enum WordWdDropPosition WordWdDropPosition;

enum WordWdNumberingRule {
	WordWdNumberingRuleRestartContinuous = '\002=\000\000',
	WordWdNumberingRuleRestartSection = '\002=\000\001',
	WordWdNumberingRuleRestartPage = '\002=\000\002'
};
typedef enum WordWdNumberingRule WordWdNumberingRule;

enum WordWdFootnoteLocation {
	WordWdFootnoteLocationBottomOfPage = '\002>\000\000',
	WordWdFootnoteLocationBeneathText = '\002>\000\001'
};
typedef enum WordWdFootnoteLocation WordWdFootnoteLocation;

enum WordWdEndnoteLocation {
	WordWdEndnoteLocationEnd_of_section = '\002\?\000\000',
	WordWdEndnoteLocationEnd_of_document = '\002\?\000\001'
};
typedef enum WordWdEndnoteLocation WordWdEndnoteLocation;

enum WordWdSortSeparator {
	WordWdSortSeparatorSortSeparateByTabs = '\002@\000\000',
	WordWdSortSeparatorSortSeparateByCommas = '\002@\000\001',
	WordWdSortSeparatorSortSeparateByDefaultTableSeparator = '\002@\000\002'
};
typedef enum WordWdSortSeparator WordWdSortSeparator;

enum WordWdTableFieldSeparator {
	WordWdTableFieldSeparatorSeparateByParagraphs = '\002A\000\000',
	WordWdTableFieldSeparatorSeparateByTabs = '\002A\000\001',
	WordWdTableFieldSeparatorSeparateByCommas = '\002A\000\002',
	WordWdTableFieldSeparatorSeparateByDefaultListSeparator = '\002A\000\003'
};
typedef enum WordWdTableFieldSeparator WordWdTableFieldSeparator;

enum WordWdSortFieldType {
	WordWdSortFieldTypeSortFieldAlphanumeric = '\002B\000\000',
	WordWdSortFieldTypeSortFieldNumeric = '\002B\000\001',
	WordWdSortFieldTypeSortFieldDate = '\002B\000\002',
	WordWdSortFieldTypeSortFieldSyllable = '\002B\000\003',
	WordWdSortFieldTypeSortFieldJapanJis = '\002B\000\004',
	WordWdSortFieldTypeSortFieldStroke = '\002B\000\005',
	WordWdSortFieldTypeSortFieldKoreaKs = '\002B\000\006'
};
typedef enum WordWdSortFieldType WordWdSortFieldType;

enum WordWdSortOrder {
	WordWdSortOrderSortOrderAscending = '\002C\000\000',
	WordWdSortOrderSortOrderDescending = '\002C\000\001'
};
typedef enum WordWdSortOrder WordWdSortOrder;

enum WordWdTableFormat {
	WordWdTableFormatTableFormatNone = '\002D\000\000',
	WordWdTableFormatTableFormatSimple1 = '\002D\000\001',
	WordWdTableFormatTableFormatSimple2 = '\002D\000\002',
	WordWdTableFormatTableFormatSimple3 = '\002D\000\003',
	WordWdTableFormatTableFormatClassic1 = '\002D\000\004',
	WordWdTableFormatTableFormatClassic2 = '\002D\000\005',
	WordWdTableFormatTableFormatClassic3 = '\002D\000\006',
	WordWdTableFormatTableFormatClassic4 = '\002D\000\007',
	WordWdTableFormatTableFormatColorful1 = '\002D\000\010',
	WordWdTableFormatTableFormatColorful2 = '\002D\000\011',
	WordWdTableFormatTableFormatColorful3 = '\002D\000\012',
	WordWdTableFormatTableFormatColumns1 = '\002D\000\013',
	WordWdTableFormatTableFormatColumns2 = '\002D\000\014',
	WordWdTableFormatTableFormatColumns3 = '\002D\000\015',
	WordWdTableFormatTableFormatColumns4 = '\002D\000\016',
	WordWdTableFormatTableFormatColumns5 = '\002D\000\017',
	WordWdTableFormatTableFormatGrid1 = '\002D\000\020',
	WordWdTableFormatTableFormatGrid2 = '\002D\000\021',
	WordWdTableFormatTableFormatGrid3 = '\002D\000\022',
	WordWdTableFormatTableFormatGrid4 = '\002D\000\023',
	WordWdTableFormatTableFormatGrid5 = '\002D\000\024',
	WordWdTableFormatTableFormatGrid6 = '\002D\000\025',
	WordWdTableFormatTableFormatGrid7 = '\002D\000\026',
	WordWdTableFormatTableFormatGrid8 = '\002D\000\027',
	WordWdTableFormatTableFormatList1 = '\002D\000\030',
	WordWdTableFormatTableFormatList2 = '\002D\000\031',
	WordWdTableFormatTableFormatList3 = '\002D\000\032',
	WordWdTableFormatTableFormatList4 = '\002D\000\033',
	WordWdTableFormatTableFormatList5 = '\002D\000\034',
	WordWdTableFormatTableFormatList6 = '\002D\000\035',
	WordWdTableFormatTableFormatList7 = '\002D\000\036',
	WordWdTableFormatTableFormatList8 = '\002D\000\037',
	WordWdTableFormatTableFormat3DEffects1 = '\002D\000 ',
	WordWdTableFormatTableFormat3DEffects2 = '\002D\000!',
	WordWdTableFormatTableFormat3DEffects3 = '\002D\000\"',
	WordWdTableFormatTableFormatContemporary = '\002D\000#',
	WordWdTableFormatTableFormatElegant = '\002D\000$',
	WordWdTableFormatTableFormatProfessional = '\002D\000%',
	WordWdTableFormatTableFormatSubtle1 = '\002D\000&',
	WordWdTableFormatTableFormatSubtle2 = '\002D\000\'',
	WordWdTableFormatTableFormatWeb1 = '\002D\000(',
	WordWdTableFormatTableFormatWeb2 = '\002D\000)',
	WordWdTableFormatTableFormatWeb3 = '\002D\000*'
};
typedef enum WordWdTableFormat WordWdTableFormat;

enum WordWdTableFormatApply {
	WordWdTableFormatApplyTableFormatApplyBorders = '\002E\000\001',
	WordWdTableFormatApplyTableFormatApplyShading = '\002E\000\002',
	WordWdTableFormatApplyTableFormatApplyFont = '\002E\000\004',
	WordWdTableFormatApplyTableFormatApplyColor = '\002E\000\010',
	WordWdTableFormatApplyTableFormatApplyAutoFit = '\002E\000\020',
	WordWdTableFormatApplyTableFormatApplyHeadingRows = '\002E\000 ',
	WordWdTableFormatApplyTableFormatApplyLastRow = '\002E\000@',
	WordWdTableFormatApplyTableFormatApplyFirstColumn = '\002E\000\200',
	WordWdTableFormatApplyTableFormatApplyLastColumn = '\002E\001\000'
};
typedef enum WordWdTableFormatApply WordWdTableFormatApply;

enum WordWdLanguageID {
	WordWdLanguageIDLanguageNone = '\002F\000\000',
	WordWdLanguageIDLanguageNoProofing = '\002F\004\000',
	WordWdLanguageIDAfrikaans = '\002F\0046',
	WordWdLanguageIDAlbanian = '\002F\004\034',
	WordWdLanguageIDAmharic = '\002F\004^',
	WordWdLanguageIDArabicAlgeria = '\002F\024\001',
	WordWdLanguageIDArabicBahrain = '\002F<\001',
	WordWdLanguageIDArabicEgypt = '\002F\014\001',
	WordWdLanguageIDArabicIraq = '\002F\010\001',
	WordWdLanguageIDArabicJordan = '\002F,\001',
	WordWdLanguageIDArabicKuwait = '\002F4\001',
	WordWdLanguageIDArabicLebanon = '\002F0\001',
	WordWdLanguageIDArabicLibya = '\002F\020\001',
	WordWdLanguageIDArabicMorocco = '\002F\030\001',
	WordWdLanguageIDArabicOman = '\002F \001',
	WordWdLanguageIDArabicQatar = '\002F@\001',
	WordWdLanguageIDArabic = '\002F\004\001',
	WordWdLanguageIDArabicSyria = '\002F(\001',
	WordWdLanguageIDArabicTunisia = '\002F\034\001',
	WordWdLanguageIDArabicUae = '\002F8\001',
	WordWdLanguageIDArabicYemen = '\002F$\001',
	WordWdLanguageIDArmenian = '\002F\004+',
	WordWdLanguageIDAssamese = '\002F\004M',
	WordWdLanguageIDAzeriCyrillic = '\002F\010,',
	WordWdLanguageIDAzeriLatin = '\002F\004,',
	WordWdLanguageIDBasque = '\002F\004-',
	WordWdLanguageIDByelorussian = '\002F\004#',
	WordWdLanguageIDBengali = '\002F\004E',
	WordWdLanguageIDBulgarian = '\002F\004\002',
	WordWdLanguageIDBurmese = '\002F\004U',
	WordWdLanguageIDCatalan = '\002F\004\003',
	WordWdLanguageIDCherokee = '\002F\004\\',
	WordWdLanguageIDChineseHongKong = '\002F\014\004',
	WordWdLanguageIDChineseMacao = '\002F\024\004',
	WordWdLanguageIDSimplifiedChinese = '\002F\010\004',
	WordWdLanguageIDChineseSingapore = '\002F\020\004',
	WordWdLanguageIDTraditionalChinese = '\002F\004\004',
	WordWdLanguageIDCroatian = '\002F\004\032',
	WordWdLanguageIDCzech = '\002F\004\005',
	WordWdLanguageIDDanish = '\002F\004\006',
	WordWdLanguageIDDivehi = '\002F\004e',
	WordWdLanguageIDBelgianDutch = '\002F\010\023',
	WordWdLanguageIDDutch = '\002F\004\023',
	WordWdLanguageIDEdo = '\002F\004f',
	WordWdLanguageIDEnglishAus = '\002F\014\011',
	WordWdLanguageIDEnglishBelize = '\002F(\011',
	WordWdLanguageIDEnglishCanadian = '\002F\020\011',
	WordWdLanguageIDEnglishCaribbean = '\002F$\011',
	WordWdLanguageIDEnglishIreland = '\002F\030\011',
	WordWdLanguageIDEnglishJamaica = '\002F \011',
	WordWdLanguageIDEnglishNewZealand = '\002F\024\011',
	WordWdLanguageIDEnglishPhilippines = '\002F4\011',
	WordWdLanguageIDEnglishSouthAfrica = '\002F\034\011',
	WordWdLanguageIDEnglishTrinidadTobago = '\002F,\011',
	WordWdLanguageIDEnglishUk = '\002F\010\011',
	WordWdLanguageIDEnglishUs = '\002F\004\011',
	WordWdLanguageIDEnglishZimbabwe = '\002F0\011',
	WordWdLanguageIDEnglishIndonesia = '\002F8\011',
	WordWdLanguageIDEstonian = '\002F\004%',
	WordWdLanguageIDFaeroese = '\002F\0048',
	WordWdLanguageIDPersian = '\002F\004)',
	WordWdLanguageIDFilipino = '\002F\004d',
	WordWdLanguageIDFinnish = '\002F\004\013',
	WordWdLanguageIDFulfulde = '\002F\004g',
	WordWdLanguageIDBelgianFrench = '\002F\010\014',
	WordWdLanguageIDFrenchCameroon = '\002F,\014',
	WordWdLanguageIDFrenchCanadian = '\002F\014\014',
	WordWdLanguageIDFrenchCotedIvoire = '\002F0\014',
	WordWdLanguageIDFrench = '\002F\004\014',
	WordWdLanguageIDFrenchLuxembourg = '\002F\024\014',
	WordWdLanguageIDFrenchMali = '\002F4\014',
	WordWdLanguageIDFrenchMonaco = '\002F\030\014',
	WordWdLanguageIDFrenchReunion = '\002F \014',
	WordWdLanguageIDFrenchSenegal = '\002F(\014',
	WordWdLanguageIDFrenchMorocco = '\002F8\014',
	WordWdLanguageIDFrenchHaiti = '\002F<\014',
	WordWdLanguageIDSwissFrench = '\002F\020\014',
	WordWdLanguageIDFrenchWestIndies = '\002F\034\014',
	WordWdLanguageIDFrenchCongo = '\002F$\014',
	WordWdLanguageIDFrisianNetherlands = '\002F\004b',
	WordWdLanguageIDGaelicIreland = '\002F\010<',
	WordWdLanguageIDGaelicScotland = '\002F\004<',
	WordWdLanguageIDGalician = '\002F\004V',
	WordWdLanguageIDGeorgian = '\002F\0047',
	WordWdLanguageIDGermanAustria = '\002F\014\007',
	WordWdLanguageIDGerman = '\002F\004\007',
	WordWdLanguageIDGermanLiechtenstein = '\002F\024\007',
	WordWdLanguageIDGermanLuxembourg = '\002F\020\007',
	WordWdLanguageIDSwissGerman = '\002F\010\007',
	WordWdLanguageIDGreek = '\002F\004\010',
	WordWdLanguageIDGuarani = '\002F\004t',
	WordWdLanguageIDGujarati = '\002F\004G',
	WordWdLanguageIDHausa = '\002F\004h',
	WordWdLanguageIDHawaiian = '\002F\004u',
	WordWdLanguageIDHebrew = '\002F\004\015',
	WordWdLanguageIDHindi = '\002F\0049',
	WordWdLanguageIDHungarian = '\002F\004\016',
	WordWdLanguageIDIbibio = '\002F\004i',
	WordWdLanguageIDIcelandic = '\002F\004\017',
	WordWdLanguageIDIgbo = '\002F\004p',
	WordWdLanguageIDIndonesian = '\002F\004!',
	WordWdLanguageIDInuktitut = '\002F\004]',
	WordWdLanguageIDItalian = '\002F\004\020',
	WordWdLanguageIDSwissItalian = '\002F\010\020',
	WordWdLanguageIDJapanese = '\002F\004\021',
	WordWdLanguageIDKannada = '\002F\004K',
	WordWdLanguageIDKanuri = '\002F\004q',
	WordWdLanguageIDKashmiri = '\002F\004`',
	WordWdLanguageIDKazakh = '\002F\004\?',
	WordWdLanguageIDKhmer = '\002F\004S',
	WordWdLanguageIDKirghiz = '\002F\004@',
	WordWdLanguageIDKonkani = '\002F\004W',
	WordWdLanguageIDKorean = '\002F\004\022',
	WordWdLanguageIDKyrgyz = '\002F\004@',
	WordWdLanguageIDLao = '\002F\004T',
	WordWdLanguageIDLatin = '\002F\004v',
	WordWdLanguageIDLatvian = '\002F\004&',
	WordWdLanguageIDLithuanian = '\002F\004\'',
	WordWdLanguageIDMacedonian = '\002F\004/',
	WordWdLanguageIDMalaysian = '\002F\004>',
	WordWdLanguageIDMalayBruneiDarussalam = '\002F\010>',
	WordWdLanguageIDMalayalam = '\002F\004L',
	WordWdLanguageIDMaltese = '\002F\004:',
	WordWdLanguageIDManipuri = '\002F\004X',
	WordWdLanguageIDMarathi = '\002F\004N',
	WordWdLanguageIDMongolian = '\002F\004P',
	WordWdLanguageIDNepali = '\002F\004a',
	WordWdLanguageIDNorwegianBokmol = '\002F\004\024',
	WordWdLanguageIDNorwegianNynorsk = '\002F\010\024',
	WordWdLanguageIDOriya = '\002F\004H',
	WordWdLanguageIDOromo = '\002F\004r',
	WordWdLanguageIDPashto = '\002F\004c',
	WordWdLanguageIDPolish = '\002F\004\025',
	WordWdLanguageIDPortugueseBrazil = '\002F\004\026',
	WordWdLanguageIDPortuguese = '\002F\010\026',
	WordWdLanguageIDPunjabi = '\002F\004F',
	WordWdLanguageIDRhaetoRomanic = '\002F\004\027',
	WordWdLanguageIDRomanianMoldova = '\002F\010\030',
	WordWdLanguageIDRomanian = '\002F\004\030',
	WordWdLanguageIDRussianMoldova = '\002F\010\031',
	WordWdLanguageIDRussian = '\002F\004\031',
	WordWdLanguageIDSamiLappish = '\002F\004;',
	WordWdLanguageIDSanskrit = '\002F\004O',
	WordWdLanguageIDSerbianCyrillic = '\002F\014\032',
	WordWdLanguageIDSerbianLatin = '\002F\010\032',
	WordWdLanguageIDSinhalese = '\002F\004[',
	WordWdLanguageIDSindhi = '\002F\004Y',
	WordWdLanguageIDSindhiPakistan = '\002F\010Y',
	WordWdLanguageIDSlovak = '\002F\004\033',
	WordWdLanguageIDSlovenian = '\002F\004$',
	WordWdLanguageIDSomali = '\002F\004w',
	WordWdLanguageIDSorbian = '\002F\004.',
	WordWdLanguageIDSpanishArgentina = '\002F,\012',
	WordWdLanguageIDSpanishBolivia = '\002F@\012',
	WordWdLanguageIDSpanishChile = '\002F4\012',
	WordWdLanguageIDSpanishColombia = '\002F$\012',
	WordWdLanguageIDSpanishCostaRica = '\002F\024\012',
	WordWdLanguageIDSpanishDominicanRepublic = '\002F\034\012',
	WordWdLanguageIDSpanishEcuador = '\002F0\012',
	WordWdLanguageIDSpanishElSalvador = '\002FD\012',
	WordWdLanguageIDSpanishGuatemala = '\002F\020\012',
	WordWdLanguageIDSpanishHonduras = '\002FH\012',
	WordWdLanguageIDMexicanSpanish = '\002F\010\012',
	WordWdLanguageIDSpanishNicaragua = '\002FL\012',
	WordWdLanguageIDSpanishPanama = '\002F\030\012',
	WordWdLanguageIDSpanishParaguay = '\002F<\012',
	WordWdLanguageIDSpanishPeru = '\002F(\012',
	WordWdLanguageIDSpanishPuertoRico = '\002FP\012',
	WordWdLanguageIDSpanishModernSort = '\002F\014\012',
	WordWdLanguageIDSpanish = '\002F\004\012',
	WordWdLanguageIDSpanishUruguay = '\002F8\012',
	WordWdLanguageIDSpanishVenezuela = '\002F \012',
	WordWdLanguageIDSesotho = '\002F\0040',
	WordWdLanguageIDSutu = '\002F\0040',
	WordWdLanguageIDSwahili = '\002F\004A',
	WordWdLanguageIDSwedishFinland = '\002F\010\035',
	WordWdLanguageIDSwedish = '\002F\004\035',
	WordWdLanguageIDSyriac = '\002F\004Z',
	WordWdLanguageIDTajik = '\002F\004(',
	WordWdLanguageIDTamazight = '\002F\004_',
	WordWdLanguageIDTamazightLatin = '\002F\010_',
	WordWdLanguageIDTamil = '\002F\004I',
	WordWdLanguageIDTatar = '\002F\004D',
	WordWdLanguageIDTelugu = '\002F\004J',
	WordWdLanguageIDThai = '\002F\004\036',
	WordWdLanguageIDTibetan = '\002F\004Q',
	WordWdLanguageIDTigrignaEthiopic = '\002F\004s',
	WordWdLanguageIDTigrignaEritrea = '\002F\010s',
	WordWdLanguageIDTsonga = '\002F\0041',
	WordWdLanguageIDTswana = '\002F\0042',
	WordWdLanguageIDTurkish = '\002F\004\037',
	WordWdLanguageIDTurkmen = '\002F\004B',
	WordWdLanguageIDUkrainian = '\002F\004\"',
	WordWdLanguageIDUrdu = '\002F\004 ',
	WordWdLanguageIDUzbekCyrillic = '\002F\010C',
	WordWdLanguageIDUzbekLatin = '\002F\004C',
	WordWdLanguageIDVenda = '\002F\0043',
	WordWdLanguageIDVietnamese = '\002F\004*',
	WordWdLanguageIDWelsh = '\002F\004R',
	WordWdLanguageIDXhosa = '\002F\0044',
	WordWdLanguageIDYi = '\002F\004x',
	WordWdLanguageIDYiddish = '\002F\004=',
	WordWdLanguageIDYoruba = '\002F\004j',
	WordWdLanguageIDZulu = '\002F\0045'
};
typedef enum WordWdLanguageID WordWdLanguageID;

enum WordWdFieldType {
	WordWdFieldTypeFieldEmpty = '\002F\377\377',
	WordWdFieldTypeFieldRef = '\002G\000\003',
	WordWdFieldTypeFieldIndexEntry = '\002G\000\004',
	WordWdFieldTypeFieldFootnoteRef = '\002G\000\005',
	WordWdFieldTypeFieldSet = '\002G\000\006',
	WordWdFieldTypeFieldIf = '\002G\000\007',
	WordWdFieldTypeFieldIndex = '\002G\000\010',
	WordWdFieldTypeFieldTocEntry = '\002G\000\011',
	WordWdFieldTypeFieldStyleRef = '\002G\000\012',
	WordWdFieldTypeFieldRefDoc = '\002G\000\013',
	WordWdFieldTypeFieldSequence = '\002G\000\014',
	WordWdFieldTypeFieldToc = '\002G\000\015',
	WordWdFieldTypeFieldInfo = '\002G\000\016',
	WordWdFieldTypeFieldTitle = '\002G\000\017',
	WordWdFieldTypeFieldSubject = '\002G\000\020',
	WordWdFieldTypeFieldAuthor = '\002G\000\021',
	WordWdFieldTypeFieldKeyWord = '\002G\000\022',
	WordWdFieldTypeFieldComments = '\002G\000\023',
	WordWdFieldTypeFieldLastSavedBy = '\002G\000\024',
	WordWdFieldTypeFieldCreateDate = '\002G\000\025',
	WordWdFieldTypeFieldSaveDate = '\002G\000\026',
	WordWdFieldTypeFieldPrintDate = '\002G\000\027',
	WordWdFieldTypeFieldRevisionNum = '\002G\000\030',
	WordWdFieldTypeFieldEditTime = '\002G\000\031',
	WordWdFieldTypeFieldNumPages = '\002G\000\032',
	WordWdFieldTypeFieldNumWords = '\002G\000\033',
	WordWdFieldTypeFieldNumChars = '\002G\000\034',
	WordWdFieldTypeFieldFileName = '\002G\000\035',
	WordWdFieldTypeFieldTemplate = '\002G\000\036',
	WordWdFieldTypeFieldDate = '\002G\000\037',
	WordWdFieldTypeFieldTime = '\002G\000 ',
	WordWdFieldTypeFieldPage = '\002G\000!',
	WordWdFieldTypeFieldExpression = '\002G\000\"',
	WordWdFieldTypeFieldQuote = '\002G\000#',
	WordWdFieldTypeFieldInclude = '\002G\000$',
	WordWdFieldTypeFieldPageRef = '\002G\000%',
	WordWdFieldTypeFieldAsk = '\002G\000&',
	WordWdFieldTypeFieldFillIn = '\002G\000\'',
	WordWdFieldTypeFieldData = '\002G\000(',
	WordWdFieldTypeFieldNext = '\002G\000)',
	WordWdFieldTypeFieldNextIf = '\002G\000*',
	WordWdFieldTypeFieldSkipIf = '\002G\000+',
	WordWdFieldTypeFieldMergeRec = '\002G\000,',
	WordWdFieldTypeFieldDde = '\002G\000-',
	WordWdFieldTypeFieldDdeauto = '\002G\000.',
	WordWdFieldTypeFieldGlossary = '\002G\000/',
	WordWdFieldTypeFieldPrint = '\002G\0000',
	WordWdFieldTypeFieldFormula = '\002G\0001',
	WordWdFieldTypeFieldGoToButton = '\002G\0002',
	WordWdFieldTypeFieldMacroButton = '\002G\0003',
	WordWdFieldTypeFieldAutoNumOutline = '\002G\0004',
	WordWdFieldTypeFieldAutoNumLegal = '\002G\0005',
	WordWdFieldTypeFieldAutoNum = '\002G\0006',
	WordWdFieldTypeFieldImport = '\002G\0007',
	WordWdFieldTypeFieldLink = '\002G\0008',
	WordWdFieldTypeFieldSymbol = '\002G\0009',
	WordWdFieldTypeFieldEmbed = '\002G\000:',
	WordWdFieldTypeFieldMergeField = '\002G\000;',
	WordWdFieldTypeFieldUserName = '\002G\000<',
	WordWdFieldTypeFieldUserInitials = '\002G\000=',
	WordWdFieldTypeFieldUserAddress = '\002G\000>',
	WordWdFieldTypeFieldBarCode = '\002G\000\?',
	WordWdFieldTypeFieldDocVariable = '\002G\000@',
	WordWdFieldTypeFieldSection = '\002G\000A',
	WordWdFieldTypeFieldSectionPages = '\002G\000B',
	WordWdFieldTypeFieldIncludePicture = '\002G\000C',
	WordWdFieldTypeFieldIncludeText = '\002G\000D',
	WordWdFieldTypeFieldFileSize = '\002G\000E',
	WordWdFieldTypeFieldFormTextInput = '\002G\000F',
	WordWdFieldTypeFieldFormCheckBox = '\002G\000G',
	WordWdFieldTypeFieldNoteRef = '\002G\000H',
	WordWdFieldTypeFieldToa = '\002G\000I',
	WordWdFieldTypeFieldToaentry = '\002G\000J',
	WordWdFieldTypeFieldMergeSeq = '\002G\000K',
	WordWdFieldTypeFieldPrivate = '\002G\000M',
	WordWdFieldTypeFieldDatabase = '\002G\000N',
	WordWdFieldTypeFieldAutoText = '\002G\000O',
	WordWdFieldTypeFieldCompare = '\002G\000P',
	WordWdFieldTypeFieldAddin = '\002G\000Q',
	WordWdFieldTypeFieldSubscriber = '\002G\000R',
	WordWdFieldTypeFieldFormDropDown = '\002G\000S',
	WordWdFieldTypeFieldAdvance = '\002G\000T',
	WordWdFieldTypeFieldDocProperty = '\002G\000U',
	WordWdFieldTypeFieldOcx = '\002G\000W',
	WordWdFieldTypeFieldHyperlink = '\002G\000X',
	WordWdFieldTypeFieldAutoTextList = '\002G\000Y',
	WordWdFieldTypeFieldListnum = '\002G\000Z',
	WordWdFieldTypeFieldHtmlactiveX = '\002G\000[',
	WordWdFieldTypeFieldBidiOutline = '\002G\000\\',
	WordWdFieldTypeFieldAddressBlock = '\002G\000]',
	WordWdFieldTypeFieldGreetingLine = '\002G\000^',
	WordWdFieldTypeFieldShape = '\002G\000_',
	WordWdFieldTypeFieldCitation = '\002G\000`',
	WordWdFieldTypeFieldBibliography = '\002G\000a',
	WordWdFieldTypeFieldMergeBarcode = '\002G\000b',
	WordWdFieldTypeFieldDisplayBarcode = '\002G\000c'
};
typedef enum WordWdFieldType WordWdFieldType;

enum WordWdBuiltinStyle {
	WordWdBuiltinStyleStyleNormal = '\002G\377\377',
	WordWdBuiltinStyleStyleHeading1 = '\002G\377\376',
	WordWdBuiltinStyleStyleHeading2 = '\002G\377\375',
	WordWdBuiltinStyleStyleHeading3 = '\002G\377\374',
	WordWdBuiltinStyleStyleHeading4 = '\002G\377\373',
	WordWdBuiltinStyleStyleHeading5 = '\002G\377\372',
	WordWdBuiltinStyleStyleHeading6 = '\002G\377\371',
	WordWdBuiltinStyleStyleHeading7 = '\002G\377\370',
	WordWdBuiltinStyleStyleHeading8 = '\002G\377\367',
	WordWdBuiltinStyleStyleHeading9 = '\002G\377\366',
	WordWdBuiltinStyleStyleIndex1 = '\002G\377\365',
	WordWdBuiltinStyleStyleIndex2 = '\002G\377\364',
	WordWdBuiltinStyleStyleIndex3 = '\002G\377\363',
	WordWdBuiltinStyleStyleIndex4 = '\002G\377\362',
	WordWdBuiltinStyleStyleIndex5 = '\002G\377\361',
	WordWdBuiltinStyleStyleIndex6 = '\002G\377\360',
	WordWdBuiltinStyleStyleIndex7 = '\002G\377\357',
	WordWdBuiltinStyleStyleIndex8 = '\002G\377\356',
	WordWdBuiltinStyleStyleIndex9 = '\002G\377\355',
	WordWdBuiltinStyleStyleToc1 = '\002G\377\354',
	WordWdBuiltinStyleStyleToc2 = '\002G\377\353',
	WordWdBuiltinStyleStyleToc3 = '\002G\377\352',
	WordWdBuiltinStyleStyleToc4 = '\002G\377\351',
	WordWdBuiltinStyleStyleToc5 = '\002G\377\350',
	WordWdBuiltinStyleStyleToc6 = '\002G\377\347',
	WordWdBuiltinStyleStyleToc7 = '\002G\377\346',
	WordWdBuiltinStyleStyleToc8 = '\002G\377\345',
	WordWdBuiltinStyleStyleToc9 = '\002G\377\344',
	WordWdBuiltinStyleStyleNormalIndent = '\002G\377\343',
	WordWdBuiltinStyleStyleFootnoteText = '\002G\377\342',
	WordWdBuiltinStyleStyleCommentText = '\002G\377\341',
	WordWdBuiltinStyleStyleHeader = '\002G\377\340',
	WordWdBuiltinStyleStyleFooter = '\002G\377\337',
	WordWdBuiltinStyleStyleIndexHeading = '\002G\377\336',
	WordWdBuiltinStyleStyleCaption = '\002G\377\335',
	WordWdBuiltinStyleStyleTableOfFigures = '\002G\377\334',
	WordWdBuiltinStyleStyleEnvelopeAddress = '\002G\377\333',
	WordWdBuiltinStyleStyleEnvelopeReturn = '\002G\377\332',
	WordWdBuiltinStyleStyleFootnoteReference = '\002G\377\331',
	WordWdBuiltinStyleStyleCommentReference = '\002G\377\330',
	WordWdBuiltinStyleStyleLineNumber = '\002G\377\327',
	WordWdBuiltinStyleStylePageNumber = '\002G\377\326',
	WordWdBuiltinStyleStyleEndnoteReference = '\002G\377\325',
	WordWdBuiltinStyleStyleEndnoteText = '\002G\377\324',
	WordWdBuiltinStyleStyleTableOfAuthorities = '\002G\377\323',
	WordWdBuiltinStyleStyleMacroText = '\002G\377\322',
	WordWdBuiltinStyleStyleToaHeading = '\002G\377\321',
	WordWdBuiltinStyleStyleList = '\002G\377\320',
	WordWdBuiltinStyleStyleListBullet = '\002G\377\317',
	WordWdBuiltinStyleStyleListNumber = '\002G\377\316',
	WordWdBuiltinStyleStyleList2 = '\002G\377\315',
	WordWdBuiltinStyleStyleList3 = '\002G\377\314',
	WordWdBuiltinStyleStyleList4 = '\002G\377\313',
	WordWdBuiltinStyleStyleList5 = '\002G\377\312',
	WordWdBuiltinStyleStyleListBullet2 = '\002G\377\311',
	WordWdBuiltinStyleStyleListBullet3 = '\002G\377\310',
	WordWdBuiltinStyleStyleListBullet4 = '\002G\377\307',
	WordWdBuiltinStyleStyleListBullet5 = '\002G\377\306',
	WordWdBuiltinStyleStyleListNumber2 = '\002G\377\305',
	WordWdBuiltinStyleStyleListNumber3 = '\002G\377\304',
	WordWdBuiltinStyleStyleListNumber4 = '\002G\377\303',
	WordWdBuiltinStyleStyleListNumber5 = '\002G\377\302',
	WordWdBuiltinStyleStyleTitle = '\002G\377\301',
	WordWdBuiltinStyleStyleClosing = '\002G\377\300',
	WordWdBuiltinStyleStyleSignature = '\002G\377\277',
	WordWdBuiltinStyleStyleDefaultParagraphFont = '\002G\377\276',
	WordWdBuiltinStyleStyleBodyText = '\002G\377\275',
	WordWdBuiltinStyleStyleBodyTextIndent = '\002G\377\274',
	WordWdBuiltinStyleStyleListContinue = '\002G\377\273',
	WordWdBuiltinStyleStyleListContinue2 = '\002G\377\272',
	WordWdBuiltinStyleStyleListContinue3 = '\002G\377\271',
	WordWdBuiltinStyleStyleListContinue4 = '\002G\377\270',
	WordWdBuiltinStyleStyleListContinue5 = '\002G\377\267',
	WordWdBuiltinStyleStyleMessageHeader = '\002G\377\266',
	WordWdBuiltinStyleStyleSubtitle = '\002G\377\265',
	WordWdBuiltinStyleStyleSalutation = '\002G\377\264',
	WordWdBuiltinStyleStyleDate = '\002G\377\263',
	WordWdBuiltinStyleStyleBodyTextFirstIndent = '\002G\377\262',
	WordWdBuiltinStyleStyleBodyTextFirstIndent2 = '\002G\377\261',
	WordWdBuiltinStyleStyleNoteHeading = '\002G\377\260',
	WordWdBuiltinStyleStyleBodyText2 = '\002G\377\257',
	WordWdBuiltinStyleStyleBodyText3 = '\002G\377\256',
	WordWdBuiltinStyleStyleBodyTextIndent2 = '\002G\377\255',
	WordWdBuiltinStyleStyleBodyTextIndent3 = '\002G\377\254',
	WordWdBuiltinStyleStyleBlockQuotation = '\002G\377\253',
	WordWdBuiltinStyleStyleHyperlink = '\002G\377\252',
	WordWdBuiltinStyleStyleHyperlinkFollowed = '\002G\377\251',
	WordWdBuiltinStyleStyleStrong = '\002G\377\250',
	WordWdBuiltinStyleStyleEmphasis = '\002G\377\247',
	WordWdBuiltinStyleStyleNavPane = '\002G\377\246',
	WordWdBuiltinStyleStylePlainText = '\002G\377\245',
	WordWdBuiltinStyleStyleHtmlNormal = '\002G\377\241',
	WordWdBuiltinStyleStyleHtmlAcronym = '\002G\377\240',
	WordWdBuiltinStyleStyleHtmlAddress = '\002G\377\237',
	WordWdBuiltinStyleStyleHtmlCite = '\002G\377\236',
	WordWdBuiltinStyleStyleHtmlCode = '\002G\377\235',
	WordWdBuiltinStyleStyleHtmlDefine = '\002G\377\234',
	WordWdBuiltinStyleStyleHtmlKeyboard = '\002G\377\233',
	WordWdBuiltinStyleStyleHtmlPre = '\002G\377\232',
	WordWdBuiltinStyleStyleHtmlSamp = '\002G\377\231',
	WordWdBuiltinStyleStyleHtmlTt = '\002G\377\230',
	WordWdBuiltinStyleStyleHtmlVar = '\002G\377\227',
	WordWdBuiltinStyleStyleNormalTable = '\002G\377\226',
	WordWdBuiltinStyleStyleNormalObject = '\002G\377b',
	WordWdBuiltinStyleStyleTableLightShading = '\002G\377a',
	WordWdBuiltinStyleStyleTableLightList = '\002G\377`',
	WordWdBuiltinStyleStyleTableLightGrid = '\002G\377_',
	WordWdBuiltinStyleStyleTableMediumShading1 = '\002G\377^',
	WordWdBuiltinStyleStyleTableMediumShading2 = '\002G\377]',
	WordWdBuiltinStyleStyleTableMediumList1 = '\002G\377\\',
	WordWdBuiltinStyleStyleTableMediumList2 = '\002G\377[',
	WordWdBuiltinStyleStyleTableMediumGrid1 = '\002G\377Z',
	WordWdBuiltinStyleStyleTableMediumGrid2 = '\002G\377Y',
	WordWdBuiltinStyleStyleTableMediumGrid3 = '\002G\377X',
	WordWdBuiltinStyleStyleTableDarkList = '\002G\377W',
	WordWdBuiltinStyleStyleTableColorfulShading = '\002G\377V',
	WordWdBuiltinStyleStyleTableColorfulList = '\002G\377U',
	WordWdBuiltinStyleStyleTableColorfulGrid = '\002G\377T',
	WordWdBuiltinStyleStyleTableLightShadingAccent1 = '\002G\377S',
	WordWdBuiltinStyleStyleTableLightListAccent1 = '\002G\377R',
	WordWdBuiltinStyleStyleTableLightGridAccent1 = '\002G\377Q',
	WordWdBuiltinStyleStyleTableMediumShading1Accent1 = '\002G\377P',
	WordWdBuiltinStyleStyleTableMediumShading2Accent1 = '\002G\377O',
	WordWdBuiltinStyleStyleTableMediumList1Accent1 = '\002G\377N',
	WordWdBuiltinStyleStyleListParagraph = '\002G\377L',
	WordWdBuiltinStyleStyleQuote = '\002G\377K',
	WordWdBuiltinStyleStyleIntenseQuote = '\002G\377J',
	WordWdBuiltinStyleStyleSubtleEmphasis = '\002G\376\373',
	WordWdBuiltinStyleStyleIntenseEmphasis = '\002G\376\372',
	WordWdBuiltinStyleStyleSubtleReference = '\002G\376\371',
	WordWdBuiltinStyleStyleIntenseReference = '\002G\376\370',
	WordWdBuiltinStyleStyleBookTitle = '\002G\376\367',
	WordWdBuiltinStyleStyleBibliography = '\002G\376\366',
	WordWdBuiltinStyleStyleTocHeading = '\002G\376\365'
};
typedef enum WordWdBuiltinStyle WordWdBuiltinStyle;

enum WordWdWordDialogTab {
	WordWdWordDialogTabDialogFileDocumentLayoutTabMargins = '\002I\000\017',
	WordWdWordDialogTabDialogFileDocumentLayoutTabLayout = '\002I\000\020',
	WordWdWordDialogTabDialogFilePageSetupTabMargins = '\002I\000\021',
	WordWdWordDialogTabDialogFilePageSetupTabPaperSize = '\002I\000\022',
	WordWdWordDialogTabDialogFilePageSetupTabPaperSource = '\002I\000\023',
	WordWdWordDialogTabDialogFilePageSetupTabLayout = '\002I\000\024',
	WordWdWordDialogTabDialogFilePageSetupTabCharsLines = '\002I\000\025',
	WordWdWordDialogTabDialogInsertSymbolTabSymbols = '\002I\000\026',
	WordWdWordDialogTabDialogInsertSymbolTabSpecialCharacters = '\002I\000\027',
	WordWdWordDialogTabDialogNoteOptionsTabAllFootnotes = '\002I\000\030',
	WordWdWordDialogTabDialogNoteOptionsTabAllEndnotes = '\002I\000\031',
	WordWdWordDialogTabDialogInsertIndexAndTablesTabIndex = '\002I\000\032',
	WordWdWordDialogTabDialogInsertIndexAndTablesTabTableOfContents = '\002I\000\033',
	WordWdWordDialogTabDialogInsertIndexAndTablesTabTableOfFigures = '\002I\000\034',
	WordWdWordDialogTabDialogInsertIndexAndTablesTabTableOfAuthorities = '\002I\000\035',
	WordWdWordDialogTabDialogOrganizerTabStyles = '\002I\000\036',
	WordWdWordDialogTabDialogOrganizerTabAutoText = '\002I\000\037',
	WordWdWordDialogTabDialogOrganizerTabCommandBars = '\002I\000 ',
	WordWdWordDialogTabDialogOrganizerTabMacros = '\002I\000!',
	WordWdWordDialogTabDialogFormatFontTabFont = '\002I\000\"',
	WordWdWordDialogTabDialogFormatFontTabCharacterSpacing = '\002I\000#',
	WordWdWordDialogTabDialogFormatBordersAndShadingTabBorders = '\002I\000%',
	WordWdWordDialogTabDialogFormatBordersAndShadingTabPageBorder = '\002I\000&',
	WordWdWordDialogTabDialogFormatBordersAndShadingTabShading = '\002I\000\'',
	WordWdWordDialogTabDialogToolsEnvelopesAndLabelsTabEnvelopes = '\002I\000(',
	WordWdWordDialogTabDialogToolsEnvelopesAndLabelsTabLabels = '\002I\000)',
	WordWdWordDialogTabDialogFormatParagraphTabIndentsAndSpacing = '\002I\000*',
	WordWdWordDialogTabDialogFormatParagraphTabTextFlow = '\002I\000+',
	WordWdWordDialogTabDialogFormatParagraphTabTeisai = '\002I\000,',
	WordWdWordDialogTabDialogFormatDrawingObjectTabColorsAndLines = '\002I\000-',
	WordWdWordDialogTabDialogFormatDrawingObjectTabSize = '\002I\000.',
	WordWdWordDialogTabDialogFormatDrawingObjectTabPosition = '\002I\000/',
	WordWdWordDialogTabDialogFormatDrawingObjectTabWrapping = '\002I\0000',
	WordWdWordDialogTabDialogFormatDrawingObjectTabPicture = '\002I\0001',
	WordWdWordDialogTabDialogFormatDrawingObjectTabTextbox = '\002I\0002',
	WordWdWordDialogTabDialogFormatDrawingObjectTabHr = '\002I\0003',
	WordWdWordDialogTabDialogToolsAutocorrectExceptionsTabFirstLetter = '\002I\0004',
	WordWdWordDialogTabDialogToolsAutocorrectExceptionsTabInitialCaps = '\002I\0005',
	WordWdWordDialogTabDialogToolsAutocorrectExceptionsTabHangulAndAlphabet = '\002I\0006',
	WordWdWordDialogTabDialogToolsAutocorrectExceptionsTabIac = '\002I\0007',
	WordWdWordDialogTabDialogFormatBulletsAndNumberingTabBulleted = '\002I\0008',
	WordWdWordDialogTabDialogFormatBulletsAndNumberingTabNumbered = '\002I\0009',
	WordWdWordDialogTabDialogFormatBulletsAndNumberingTabOutlineNumbered = '\002I\000:',
	WordWdWordDialogTabDialogLetterWizardTabLetterFormat = '\002I\000;',
	WordWdWordDialogTabDialogLetterWizardTabRecipientInfo = '\002I\000<',
	WordWdWordDialogTabDialogLetterWizardTabOtherElements = '\002I\000=',
	WordWdWordDialogTabDialogLetterWizardTabSenderInfo = '\002I\000>',
	WordWdWordDialogTabDialogToolsAutoManagerTabAutocorrect = '\002I\000\?',
	WordWdWordDialogTabDialogToolsAutoManagerTabMathAutocorrect = '\002I\000@',
	WordWdWordDialogTabDialogToolsAutoManagerTabAutoFormatAsYouType = '\002I\000A',
	WordWdWordDialogTabDialogToolsAutoManagerTabAutoText = '\002I\000B',
	WordWdWordDialogTabDialogToolsAutoManagerTabAutoFormat = '\002I\000C',
	WordWdWordDialogTabDialogWebOptionsGeneral = '\002I\000D',
	WordWdWordDialogTabDialogWebOptionsFiles = '\002I\000E',
	WordWdWordDialogTabDialogWebOptionsPictures = '\002I\000F',
	WordWdWordDialogTabDialogWebOptionsEncoding = '\002I\000G',
	WordWdWordDialogTabDialogWebOptionsFonts = '\002I\000H',
	WordWdWordDialogTabDialogFormatDrawingObjectTabAltText = '\002I\000I'
};
typedef enum WordWdWordDialogTab WordWdWordDialogTab;

enum WordWdWordDialog {
	WordWdWordDialogDialogHelpAbout = '\002J\000\011',
	WordWdWordDialogDialogHelpWordPerfectHelp = '\002J\000\012',
	WordWdWordDialogDialogDocumentStatistics = '\002J\000N',
	WordWdWordDialogDialogFileNew = '\002J\000O',
	WordWdWordDialogDialogFileOpen = '\002J\000P',
	WordWdWordDialogDialogDataMergeOpenDataSource = '\002J\000Q',
	WordWdWordDialogDialogDataMergeOpenHeaderSource = '\002J\000R',
	WordWdWordDialogDialogFileSaveAs = '\002J\000T',
	WordWdWordDialogDialogFileSummaryInfo = '\002J\000V',
	WordWdWordDialogDialogToolsTemplates = '\002J\000W',
	WordWdWordDialogDialogFilePrint = '\002J\000X',
	WordWdWordDialogDialogFilePrintSetup = '\002J\000a',
	WordWdWordDialogDialogFileFind = '\002J\000c',
	WordWdWordDialogDialogFormatAddressFonts = '\002J\000g',
	WordWdWordDialogDialogEditPasteSpecial = '\002J\000o',
	WordWdWordDialogDialogEditFind = '\002J\000p',
	WordWdWordDialogDialogEditReplace = '\002J\000u',
	WordWdWordDialogDialogEditStyle = '\002J\000x',
	WordWdWordDialogDialogEditLinks = '\002J\000|',
	WordWdWordDialogDialogEditObject = '\002J\000}',
	WordWdWordDialogDialogTableToText = '\002J\000\200',
	WordWdWordDialogDialogTextToTable = '\002J\000\177',
	WordWdWordDialogDialogTableInsertTable = '\002J\000\201',
	WordWdWordDialogDialogTableInsertCells = '\002J\000\202',
	WordWdWordDialogDialogTableInsertRow = '\002J\000\203',
	WordWdWordDialogDialogTableDeleteCells = '\002J\000\205',
	WordWdWordDialogDialogTableSplitCells = '\002J\000\211',
	WordWdWordDialogDialogTableRowHeight = '\002J\000\216',
	WordWdWordDialogDialogTableColumnWidth = '\002J\000\217',
	WordWdWordDialogDialogToolsCustomize = '\002J\000\230',
	WordWdWordDialogDialogInsertBreak = '\002J\000\237',
	WordWdWordDialogDialogInsertSymbol = '\002J\000\242',
	WordWdWordDialogDialogInsertPicture = '\002J\000\243',
	WordWdWordDialogDialogInsertFile = '\002J\000\244',
	WordWdWordDialogDialogInsertDateTime = '\002J\000\245',
	WordWdWordDialogDialogInsertField = '\002J\000\246',
	WordWdWordDialogDialogInsertMergeField = '\002J\000\247',
	WordWdWordDialogDialogInsertBookmark = '\002J\000\250',
	WordWdWordDialogDialogMarkIndexEntry = '\002J\000\251',
	WordWdWordDialogDialogInsertIndex = '\002J\000\252',
	WordWdWordDialogDialogInsertTableOfContents = '\002J\000\253',
	WordWdWordDialogDialogInsertObject = '\002J\000\254',
	WordWdWordDialogDialogToolsCreateEnvelope = '\002J\000\255',
	WordWdWordDialogDialogFormatFont = '\002J\000\256',
	WordWdWordDialogDialogFormatParagraph = '\002J\000\257',
	WordWdWordDialogDialogFormatSectionLayout = '\002J\000\260',
	WordWdWordDialogDialogFormatColumns = '\002J\000\261',
	WordWdWordDialogDialogFileDocumentLayout = '\002J\000\262',
	WordWdWordDialogDialogFilePageSetup = '\002J\000\262',
	WordWdWordDialogDialogFormatTabs = '\002J\000\263',
	WordWdWordDialogDialogFormatStyle = '\002J\000\264',
	WordWdWordDialogDialogFormatDefineStyleFont = '\002J\000\265',
	WordWdWordDialogDialogFormatDefineStylePara = '\002J\000\266',
	WordWdWordDialogDialogFormatDefineStyleTabs = '\002J\000\267',
	WordWdWordDialogDialogFormatDefineStyleFrame = '\002J\000\270',
	WordWdWordDialogDialogFormatDefineStyleBorders = '\002J\000\271',
	WordWdWordDialogDialogFormatDefineStyleLang = '\002J\000\272',
	WordWdWordDialogDialogFormatPicture = '\002J\000\273',
	WordWdWordDialogDialogToolsLanguage = '\002J\000\274',
	WordWdWordDialogDialogFormatBordersAndShading = '\002J\000\275',
	WordWdWordDialogDialogFormatFrame = '\002J\000\276',
	WordWdWordDialogDialogToolsThesaurus = '\002J\000\302',
	WordWdWordDialogDialogToolsHyphenation = '\002J\000\303',
	WordWdWordDialogDialogToolsBulletsNumbers = '\002J\000\304',
	WordWdWordDialogDialogToolsHighlightChanges = '\002J\000\305',
	WordWdWordDialogDialogToolsRevisions = '\002J\000\305',
	WordWdWordDialogDialogToolsCompareDocuments = '\002J\000\306',
	WordWdWordDialogDialogTableSort = '\002J\000\307',
	WordWdWordDialogDialogToolsOptionsGeneral = '\002J\000\313',
	WordWdWordDialogDialogToolsOptionsView = '\002J\000\314',
	WordWdWordDialogDialogToolsAdvancedSettings = '\002J\000\316',
	WordWdWordDialogDialogToolsOptionsPrint = '\002J\000\320',
	WordWdWordDialogDialogToolsOptionsSave = '\002J\000\321',
	WordWdWordDialogDialogToolsOptionsSpellingAndGrammar = '\002J\000\323',
	WordWdWordDialogDialogToolsOptionsUserInfo = '\002J\000\325',
	WordWdWordDialogDialogToolsMacroRecord = '\002J\000\326',
	WordWdWordDialogDialogToolsMacro = '\002J\000\327',
	WordWdWordDialogDialogWindowActivate = '\002J\000\334',
	WordWdWordDialogDialogFormatRetAddrFonts = '\002J\000\335',
	WordWdWordDialogDialogOrganizer = '\002J\000\336',
	WordWdWordDialogDialogToolsOptionsEdit = '\002J\000\340',
	WordWdWordDialogDialogToolsOptionsFileLocations = '\002J\000\341',
	WordWdWordDialogDialogToolsWordCount = '\002J\000\344',
	WordWdWordDialogDialogControlRun = '\002J\000\353',
	WordWdWordDialogDialogInsertPageNumbers = '\002J\001&',
	WordWdWordDialogDialogFormatPageNumber = '\002J\001*',
	WordWdWordDialogDialogCopyFile = '\002J\001,',
	WordWdWordDialogDialogFormatChangeCase = '\002J\001B',
	WordWdWordDialogDialogUpdate = '\002J\001K',
	WordWdWordDialogDialogInsertDatabase = '\002J\001U',
	WordWdWordDialogDialogTableFormula = '\002J\001\\',
	WordWdWordDialogDialogFormFieldOptions = '\002J\001a',
	WordWdWordDialogDialogInsertCaption = '\002J\001e',
	WordWdWordDialogDialogInsertCaptionNumbering = '\002J\001f',
	WordWdWordDialogDialogInsertAutoCaption = '\002J\001g',
	WordWdWordDialogDialogFormFieldHelp = '\002J\001i',
	WordWdWordDialogDialogInsertCrossReference = '\002J\001o',
	WordWdWordDialogDialogInsertFootnote = '\002J\001r',
	WordWdWordDialogDialogNoteOptions = '\002J\001u',
	WordWdWordDialogDialogToolsAutoCorrect = '\002J\001z',
	WordWdWordDialogDialogToolsOptionsTrackChanges = '\002J\001\202',
	WordWdWordDialogDialogConvertObject = '\002J\001\210',
	WordWdWordDialogDialogInsertAddCaption = '\002J\001\222',
	WordWdWordDialogDialogConnect = '\002J\001\244',
	WordWdWordDialogDialogToolsCustomizeKeyboard = '\002J\001\260',
	WordWdWordDialogDialogToolsCustomizeMenus = '\002J\001\261',
	WordWdWordDialogDialogToolsMergeDocuments = '\002J\001\263',
	WordWdWordDialogDialogMarkTableOfContentsEntry = '\002J\001\272',
	WordWdWordDialogDialogFilePrintOneCopy = '\002J\001\275',
	WordWdWordDialogDialogEditFrame = '\002J\001\312',
	WordWdWordDialogDialogMarkCitation = '\002J\001\317',
	WordWdWordDialogDialogTableOfContentsOptions = '\002J\001\326',
	WordWdWordDialogDialogInsertTableOfAuthorities = '\002J\001\327',
	WordWdWordDialogDialogInsertTableOfFigures = '\002J\001\330',
	WordWdWordDialogDialogInsertIndexAndTables = '\002J\001\331',
	WordWdWordDialogDialogInsertFormField = '\002J\001\343',
	WordWdWordDialogDialogFormatDropCap = '\002J\001\350',
	WordWdWordDialogDialogToolsCreateLabels = '\002J\001\351',
	WordWdWordDialogDialogToolsProtectDocument = '\002J\001\367',
	WordWdWordDialogDialogFormatStyleGallery = '\002J\001\371',
	WordWdWordDialogDialogToolsAcceptRejectChanges = '\002J\001\372',
	WordWdWordDialogDialogHelpWordPerfectHelpOptions = '\002J\001\377',
	WordWdWordDialogDialogToolsUnprotectDocument = '\002J\002\011',
	WordWdWordDialogDialogToolsOptionsCompatibility = '\002J\002\015',
	WordWdWordDialogDialogTableOfCaptionsOptions = '\002J\002\'',
	WordWdWordDialogDialogTableAutoFormat = '\002J\0023',
	WordWdWordDialogDialogMailMergeFindRecord = '\002J\0029',
	WordWdWordDialogDialogReviewAfmtRevisions = '\002J\002:',
	WordWdWordDialogDialogViewZoom = '\002J\002A',
	WordWdWordDialogDialogToolsProtectSection = '\002J\002B',
	WordWdWordDialogDialogFontSubstitution = '\002J\002E',
	WordWdWordDialogDialogInsertSubdocument = '\002J\002G',
	WordWdWordDialogDialogNewToolbar = '\002J\002J',
	WordWdWordDialogDialogToolsEnvelopesAndLabels = '\002J\002_',
	WordWdWordDialogDialogFormatCallout = '\002J\002b',
	WordWdWordDialogDialogTableFormatCell = '\002J\002d',
	WordWdWordDialogDialogToolsCustomizeMenuBar = '\002J\002g',
	WordWdWordDialogDialogFileRoutingSlip = '\002J\002p',
	WordWdWordDialogDialogEditCategory = '\002J\002q',
	WordWdWordDialogDialogToolsManageFields = '\002J\002w',
	WordWdWordDialogDialogDrawSnapToGrid = '\002J\002y',
	WordWdWordDialogDialogDrawAlign = '\002J\002z',
	WordWdWordDialogDialogMailMergeCreateDataSource = '\002J\002\202',
	WordWdWordDialogDialogMailMergeCreateHeaderSource = '\002J\002\203',
	WordWdWordDialogDialogMailMerge = '\002J\002\244',
	WordWdWordDialogDialogMailMergeCheck = '\002J\002\245',
	WordWdWordDialogDialogMailMergeHelper = '\002J\002\250',
	WordWdWordDialogDialogMailMergeQueryOptions = '\002J\002\251',
	WordWdWordDialogDialogFileMacPageSetup = '\002J\002\255',
	WordWdWordDialogDialogListCommands = '\002J\002\323',
	WordWdWordDialogDialogEditCreatePublisher = '\002J\002\334',
	WordWdWordDialogDialogEditSubscribeTo = '\002J\002\335',
	WordWdWordDialogDialogEditPublishOptions = '\002J\002\337',
	WordWdWordDialogDialogEditSubscribeOptions = '\002J\002\340',
	WordWdWordDialogDialogFileMacCustomPageSetup = '\002J\002\341',
	WordWdWordDialogDialogToolsOptionsTypography = '\002J\002\343',
	WordWdWordDialogDialogToolsAutoCorrectExceptions = '\002J\002\372',
	WordWdWordDialogDialogToolsOptionsAutoFormatAsYouType = '\002J\003\012',
	WordWdWordDialogDialogMailMergeUseAddressBook = '\002J\003\013',
	WordWdWordDialogDialogToolsHangulHanjaConversion = '\002J\003\020',
	WordWdWordDialogDialogToolsOptionsFuzzy = '\002J\003\026',
	WordWdWordDialogDialogEditGoToOld = '\002J\003+',
	WordWdWordDialogDialogInsertNumber = '\002J\003,',
	WordWdWordDialogDialogLetterWizard = '\002J\0035',
	WordWdWordDialogDialogFormatBulletsAndNumbering = '\002J\0038',
	WordWdWordDialogDialogToolsSpellingAndGrammar = '\002J\003<',
	WordWdWordDialogDialogToolsCreateDirectory = '\002J\003A',
	WordWdWordDialogDialogTableWrapping = '\002J\003V',
	WordWdWordDialogDialogFormatTheme = '\002J\003W',
	WordWdWordDialogDialogTableProperties = '\002J\003]',
	WordWdWordDialogDialogEmailOptions = '\002J\003_',
	WordWdWordDialogDialogCreateAutoText = '\002J\003h',
	WordWdWordDialogDialogToolsAutoSummarize = '\002J\003j',
	WordWdWordDialogDialogToolsGrammarSettings = '\002J\003u',
	WordWdWordDialogDialogEditGoTo = '\002J\003\200',
	WordWdWordDialogDialogWebOptions = '\002J\003\202',
	WordWdWordDialogDialogInsertHyperlink = '\002J\003\235',
	WordWdWordDialogDialogToolsAutoManager = '\002J\003\223',
	WordWdWordDialogDialogFileVersions = '\002J\003\261',
	WordWdWordDialogDialogToolsOptionsAutoFormat = '\002J\003\277',
	WordWdWordDialogDialogFormatDrawingObject = '\002J\003\300',
	WordWdWordDialogDialogToolsOptions = '\002J\003\316',
	WordWdWordDialogDialogFitText = '\002J\003\327',
	WordWdWordDialogDialogEditAutoText = '\002J\003\331',
	WordWdWordDialogDialogPhoneticGuide = '\002J\003\332',
	WordWdWordDialogDialogToolsDictionary = '\002J\003\335',
	WordWdWordDialogDialogFileSaveVersion = '\002J\003\357',
	WordWdWordDialogDialogToolsOptionsBidi = '\002J\004\005',
	WordWdWordDialogDialogFrameSetProperties = '\002J\0042',
	WordWdWordDialogDialogTableTableOptions = '\002J\0048',
	WordWdWordDialogDialogTableCellOptions = '\002J\0049',
	WordWdWordDialogDialogSetDefault = '\002J\004F',
	WordWdWordDialogDialogTranslator = '\002J\004\204',
	WordWdWordDialogDialogHorizontalInVertical = '\002J\004\210',
	WordWdWordDialogDialogTwoLinesInOne = '\002J\004\211',
	WordWdWordDialogDialogFormatEncloseCharacters = '\002J\004\212',
	WordWdWordDialogDialogConsistencyChecker = '\002J\004a',
	WordWdWordDialogDialogToolsOptionsSmartTag = '\002J\005s',
	WordWdWordDialogDialogFormatStylesCustom = '\002J\004\340',
	WordWdWordDialogDialogLinks = '\002J\004\355',
	WordWdWordDialogDialogInsertWebComponent = '\002J\005,',
	WordWdWordDialogDialogToolsOptionsEditCopyPaste = '\002J\005L',
	WordWdWordDialogDialogToolsOptionsSecurity = '\002J\005Q',
	WordWdWordDialogDialogSearch = '\002J\005S',
	WordWdWordDialogDialogShowRepairs = '\002J\005e',
	WordWdWordDialogDialogMailMergeInsertAsk = '\002J\017\317',
	WordWdWordDialogDialogMailMergeInsertFillIn = '\002J\017\320',
	WordWdWordDialogDialogMailMergeInsertIf = '\002J\017\321',
	WordWdWordDialogDialogMailMergeInsertNextIf = '\002J\017\325',
	WordWdWordDialogDialogMailMergeInsertSet = '\002J\017\326',
	WordWdWordDialogDialogMailMergeInsertSkipIf = '\002J\017\327',
	WordWdWordDialogDialogMailMergeFieldMapping = '\002J\005\030',
	WordWdWordDialogDialogMailMergeInsertAddressBlock = '\002J\005\031',
	WordWdWordDialogDialogMailMergeInsertGreetingLine = '\002J\005\032',
	WordWdWordDialogDialogMailMergeInsertFields = '\002J\005\033',
	WordWdWordDialogDialogMailMergeRecipients = '\002J\005\034',
	WordWdWordDialogDialogMailMergeFindRecipient = '\002J\005.',
	WordWdWordDialogDialogMailMergeSetDocumentType = '\002J\005;',
	WordWdWordDialogDialogLabelOptions = '\002J\005W',
	WordWdWordDialogDialogElementAttributes = '\002J\005\264',
	WordWdWordDialogDialogSchemaLibrary = '\002J\005\211',
	WordWdWordDialogDialogPermission = '\002J\005\275',
	WordWdWordDialogDialogMyPermission = '\002J\005\235',
	WordWdWordDialogDialogOptions = '\002J\005\221',
	WordWdWordDialogDialogFormattingRestrictions = '\002J\005\223',
	WordWdWordDialogDialogSourceManager = '\002J\007\200',
	WordWdWordDialogDialogCreateSource = '\002J\007\202',
	WordWdWordDialogDialogDocumentInspector = '\002J\005\312',
	WordWdWordDialogDialogStyleManagement = '\002J\007\234',
	WordWdWordDialogDialogInsertSource = '\002J\010H',
	WordWdWordDialogDialogMathRecognizedFunctions = '\002J\010u',
	WordWdWordDialogDialogInsertPlaceholder = '\002J\011,',
	WordWdWordDialogDialogBuildingBlockOrganizer = '\002J\010\023',
	WordWdWordDialogDialogContentControlProperties = '\002J\011Z',
	WordWdWordDialogDialogCompatibilityChecker = '\002J\011\207',
	WordWdWordDialogDialogExportAsFixedFormat = '\002J\011-'
	//WordWdWordDialogDialogFileNew = '\002J\004\\'
};
typedef enum WordWdWordDialog WordWdWordDialog;

enum WordWdFieldKind {
	WordWdFieldKindFieldKindNone = '\002K\000\000',
	WordWdFieldKindFieldKindHot = '\002K\000\001',
	WordWdFieldKindFieldKindWarm = '\002K\000\002',
	WordWdFieldKindFieldKindCold = '\002K\000\003'
};
typedef enum WordWdFieldKind WordWdFieldKind;

enum WordWdTextFormFieldType {
	WordWdTextFormFieldTypeRegularText = '\002L\000\000',
	WordWdTextFormFieldTypeNumberText = '\002L\000\001',
	WordWdTextFormFieldTypeDateText = '\002L\000\002',
	WordWdTextFormFieldTypeCurrentDateText = '\002L\000\003',
	WordWdTextFormFieldTypeCurrentTimeText = '\002L\000\004',
	WordWdTextFormFieldTypeCalculationText = '\002L\000\005'
};
typedef enum WordWdTextFormFieldType WordWdTextFormFieldType;

enum WordWdChevronConvertRule {
	WordWdChevronConvertRuleNeverConvert = '\002M\000\000',
	WordWdChevronConvertRuleAlwaysConvert = '\002M\000\001',
	WordWdChevronConvertRuleAskToNotConvert = '\002M\000\002',
	WordWdChevronConvertRuleAskToConvert = '\002M\000\003'
};
typedef enum WordWdChevronConvertRule WordWdChevronConvertRule;

enum WordWdMailMergeMainDocType {
	WordWdMailMergeMainDocTypeNotAMergeDocument = '\002M\377\377',
	WordWdMailMergeMainDocTypeDocumentTypeFormLetters = '\002N\000\000',
	WordWdMailMergeMainDocTypeDocumentTypeMailingLabels = '\002N\000\001',
	WordWdMailMergeMainDocTypeDocumentTypeEnvelopes = '\002N\000\002',
	WordWdMailMergeMainDocTypeDocumentTypeCatalog = '\002N\000\003',
	WordWdMailMergeMainDocTypeEmail = '\002N\000\004',
	WordWdMailMergeMainDocTypeFax = '\002N\000\005',
	WordWdMailMergeMainDocTypeDirectory = '\002N\000\003'
};
typedef enum WordWdMailMergeMainDocType WordWdMailMergeMainDocType;

enum WordWdMailMergeState {
	WordWdMailMergeStateNormalDocument = '\002O\000\000',
	WordWdMailMergeStateMainDocumentOnly = '\002O\000\001',
	WordWdMailMergeStateMainAndDataSource = '\002O\000\002',
	WordWdMailMergeStateMainAndHeader = '\002O\000\003',
	WordWdMailMergeStateMainAndSourceAndHeader = '\002O\000\004',
	WordWdMailMergeStateDataSource = '\002O\000\005'
};
typedef enum WordWdMailMergeState WordWdMailMergeState;

enum WordWdMailMergeDestination {
	WordWdMailMergeDestinationSendToNewDocument = '\002P\000\000',
	WordWdMailMergeDestinationSendToPrinter = '\002P\000\001',
	WordWdMailMergeDestinationSendToEmail = '\002P\000\002',
	WordWdMailMergeDestinationSendToFax = '\002P\000\003'
};
typedef enum WordWdMailMergeDestination WordWdMailMergeDestination;

enum WordWdMailMergeActiveRecord {
	WordWdMailMergeActiveRecordNoActiveRecord = '\002P\377\377',
	WordWdMailMergeActiveRecordNextDataRecord = '\002P\377\376',
	WordWdMailMergeActiveRecordPreviousDataRecord = '\002P\377\375',
	WordWdMailMergeActiveRecordFirstDataRecord = '\002P\377\374',
	WordWdMailMergeActiveRecordLastDataRecord = '\002P\377\373',
	WordWdMailMergeActiveRecordFirstDataSourceRecord = '\002P\377\372',
	WordWdMailMergeActiveRecordLastDataSourceRecord = '\002P\377\371',
	WordWdMailMergeActiveRecordNextDataSourceRecord = '\002P\377\370',
	WordWdMailMergeActiveRecordPreviousDataSourceRecord = '\002P\377\367'
};
typedef enum WordWdMailMergeActiveRecord WordWdMailMergeActiveRecord;

enum WordWdMailMergeDefaultRecord {
	WordWdMailMergeDefaultRecordDefaultFirstRecord = '\000\000\000\001',
	WordWdMailMergeDefaultRecordDefaultLastRecord = '\377\377\377\360'
};
typedef enum WordWdMailMergeDefaultRecord WordWdMailMergeDefaultRecord;

enum WordWdMailMergeDataSource {
	WordWdMailMergeDataSourceNoMergeInfo = '\002R\377\377',
	WordWdMailMergeDataSourceMergeInfoFromWord = '\002S\000\000',
	WordWdMailMergeDataSourceMergeInfoFromAccessDde = '\002S\000\001',
	WordWdMailMergeDataSourceMergeInfoFromExcelDde = '\002S\000\002',
	WordWdMailMergeDataSourceMergeInfoFromMsqueryDde = '\002S\000\003',
	WordWdMailMergeDataSourceMergeInfoFromOdbc = '\002S\000\004'
};
typedef enum WordWdMailMergeDataSource WordWdMailMergeDataSource;

enum WordWdMailMergeComparison {
	WordWdMailMergeComparisonMergeIfEqual = '\002T\000\000',
	WordWdMailMergeComparisonMergeIfNotEqual = '\002T\000\001',
	WordWdMailMergeComparisonMergeIfLessThan = '\002T\000\002',
	WordWdMailMergeComparisonMergeIfGreaterThan = '\002T\000\003',
	WordWdMailMergeComparisonMergeIfLessThanOrEqual = '\002T\000\004',
	WordWdMailMergeComparisonMergeIfGreaterThanOrEqual = '\002T\000\005',
	WordWdMailMergeComparisonMergeIfIsBlank = '\002T\000\006',
	WordWdMailMergeComparisonMergeIfIsNotBlank = '\002T\000\007'
};
typedef enum WordWdMailMergeComparison WordWdMailMergeComparison;

enum WordWdBookmarkSortBy {
	WordWdBookmarkSortBySortByName = '\002U\000\000',
	WordWdBookmarkSortBySortByLocation = '\002U\000\001'
};
typedef enum WordWdBookmarkSortBy WordWdBookmarkSortBy;

enum WordWdWindowState {
	WordWdWindowStateWindowStateNormal = '\002V\000\000',
	WordWdWindowStateWindowStateMaximize = '\002V\000\001',
	WordWdWindowStateWindowStateMinimize = '\002V\000\002'
};
typedef enum WordWdWindowState WordWdWindowState;

enum WordWdPictureLinkType {
	WordWdPictureLinkTypeLinkNone = '\002W\000\000',
	WordWdPictureLinkTypeLinkDataInDoc = '\002W\000\001',
	WordWdPictureLinkTypeLinkDataOnDisk = '\002W\000\002'
};
typedef enum WordWdPictureLinkType WordWdPictureLinkType;

enum WordWdLinkType {
	WordWdLinkTypeLinkTypeOle = '\002X\000\000',
	WordWdLinkTypeLinkTypePicture = '\002X\000\001',
	WordWdLinkTypeLinkTypeText = '\002X\000\002',
	WordWdLinkTypeLinkTypeReference = '\002X\000\003',
	WordWdLinkTypeLinkTypeInclude = '\002X\000\004',
	WordWdLinkTypeLinkTypeImport = '\002X\000\005',
	WordWdLinkTypeLinkTypeDde = '\002X\000\006',
	WordWdLinkTypeLinkTypeDdeauto = '\002X\000\007',
	WordWdLinkTypeLinkTypeChart = '\002X\000\010'
};
typedef enum WordWdLinkType WordWdLinkType;

enum WordWdWindowType {
	WordWdWindowTypeWindowDocument = '\002Y\000\000',
	WordWdWindowTypeWindowTemplate = '\002Y\000\001'
};
typedef enum WordWdWindowType WordWdWindowType;

enum WordWdViewType {
	WordWdViewTypeNormalView = '\002Z\000\001',
	WordWdViewTypeDraftView = '\002Z\000\001',
	WordWdViewTypeOutlineView = '\002Z\000\002',
	WordWdViewTypePrintView = '\002Z\000\003',
	WordWdViewTypePageView = '\002Z\000\003',
	WordWdViewTypePrintPreviewView = '\002Z\000\004',
	WordWdViewTypeMasterView = '\002Z\000\005',
	WordWdViewTypeWebView = '\002Z\000\006',
	WordWdViewTypeOnlineView = '\002Z\000\006',
	WordWdViewTypeReadingView = '\002Z\000\007',
	WordWdViewTypeConflictView = '\002Z\000\010'
};
typedef enum WordWdViewType WordWdViewType;

enum WordWdSeekView {
	WordWdSeekViewSeekMainDocument = '\002[\000\000',
	WordWdSeekViewSeekPrimaryHeader = '\002[\000\001',
	WordWdSeekViewSeekFirstPageHeader = '\002[\000\002',
	WordWdSeekViewSeekEvenPagesHeader = '\002[\000\003',
	WordWdSeekViewSeekPrimaryFooter = '\002[\000\004',
	WordWdSeekViewSeekFirstPageFooter = '\002[\000\005',
	WordWdSeekViewSeekEvenPagesFooter = '\002[\000\006',
	WordWdSeekViewSeekFootnotes = '\002[\000\007',
	WordWdSeekViewSeekEndnotes = '\002[\000\010',
	WordWdSeekViewSeekCurrentPageHeader = '\002[\000\011',
	WordWdSeekViewSeekCurrentPageFooter = '\002[\000\012'
};
typedef enum WordWdSeekView WordWdSeekView;

enum WordWdSpecialPane {
	WordWdSpecialPanePaneNone = '\002\\\000\000',
	WordWdSpecialPanePanePrimaryHeader = '\002\\\000\001',
	WordWdSpecialPanePaneFirstPageHeader = '\002\\\000\002',
	WordWdSpecialPanePaneEvenPagesHeader = '\002\\\000\003',
	WordWdSpecialPanePanePrimaryFooter = '\002\\\000\004',
	WordWdSpecialPanePaneFirstPageFooter = '\002\\\000\005',
	WordWdSpecialPanePaneEvenPagesFooter = '\002\\\000\006',
	WordWdSpecialPanePaneFootnotes = '\002\\\000\007',
	WordWdSpecialPanePaneEndnotes = '\002\\\000\010',
	WordWdSpecialPanePaneFootnoteContinuationNotice = '\002\\\000\011',
	WordWdSpecialPanePaneFootnoteContinuationSeparator = '\002\\\000\012',
	WordWdSpecialPanePaneFootnoteSeparator = '\002\\\000\013',
	WordWdSpecialPanePaneEndnoteContinuationNotice = '\002\\\000\014',
	WordWdSpecialPanePaneEndnoteContinuationSeparator = '\002\\\000\015',
	WordWdSpecialPanePaneEndnoteSeparator = '\002\\\000\016',
	WordWdSpecialPanePaneComments = '\002\\\000\017',
	WordWdSpecialPanePaneCurrentPageHeader = '\002\\\000\020',
	WordWdSpecialPanePaneCurrentPageFooter = '\002\\\000\021',
	WordWdSpecialPanePaneRevisions = '\002\\\000\022',
	WordWdSpecialPanePaneRevisionsHoriz = '\002\\\000\023',
	WordWdSpecialPanePaneRevisionsVert = '\002\\\000\024'
};
typedef enum WordWdSpecialPane WordWdSpecialPane;

enum WordWdPageFit {
	WordWdPageFitPageFitNone = '\002]\000\000',
	WordWdPageFitPageFitFullPage = '\002]\000\001',
	WordWdPageFitPageFitBestFit = '\002]\000\002',
	WordWdPageFitPageFitTextFit = '\002]\000\003'
};
typedef enum WordWdPageFit WordWdPageFit;

enum WordWdBrowseTarget {
	WordWdBrowseTargetBrowsePage = '\002^\000\001',
	WordWdBrowseTargetBrowseSection = '\002^\000\002',
	WordWdBrowseTargetBrowseComment = '\002^\000\003',
	WordWdBrowseTargetBrowseFootnote = '\002^\000\004',
	WordWdBrowseTargetBrowseEndnote = '\002^\000\005',
	WordWdBrowseTargetBrowseField = '\002^\000\006',
	WordWdBrowseTargetBrowseTable = '\002^\000\007',
	WordWdBrowseTargetBrowseGraphic = '\002^\000\010',
	WordWdBrowseTargetBrowseHeading = '\002^\000\011',
	WordWdBrowseTargetBrowseEdit = '\002^\000\012',
	WordWdBrowseTargetBrowseFind = '\002^\000\013',
	WordWdBrowseTargetBrowseGoTo = '\002^\000\014'
};
typedef enum WordWdBrowseTarget WordWdBrowseTarget;

enum WordWdPaperTray {
	WordWdPaperTrayPrinterDefaultBin = '\002_\000\000',
	WordWdPaperTrayPrinterUpperBin = '\002_\000\001',
	WordWdPaperTrayPrinterOnlyBin = '\002_\000\001',
	WordWdPaperTrayPrinterLowerBin = '\002_\000\002',
	WordWdPaperTrayPrinterMiddleBin = '\002_\000\003',
	WordWdPaperTrayPrinterManualFeed = '\002_\000\004',
	WordWdPaperTrayPrinterEnvelopeFeed = '\002_\000\005',
	WordWdPaperTrayPrinterManualEnvelopeFeed = '\002_\000\006',
	WordWdPaperTrayPrinterAutomaticSheetFeed = '\002_\000\007',
	WordWdPaperTrayPrinterTractorFeed = '\002_\000\010',
	WordWdPaperTrayPrinterSmallFormatBin = '\002_\000\011',
	WordWdPaperTrayPrinterLargeFormatBin = '\002_\000\012',
	WordWdPaperTrayPrinterLargeCapacityBin = '\002_\000\013',
	WordWdPaperTrayPrinterPaperCassette = '\002_\000\016',
	WordWdPaperTrayPrinterFormSource = '\002_\000\017'
};
typedef enum WordWdPaperTray WordWdPaperTray;

enum WordWdOrientation {
	WordWdOrientationOrientPortrait = '\002`\000\000',
	WordWdOrientationOrientLandscape = '\002`\000\001'
};
typedef enum WordWdOrientation WordWdOrientation;

enum WordWdSelectionType {
	WordWdSelectionTypeNoSelection = '\002a\000\000',
	WordWdSelectionTypeSelectionIp = '\002a\000\001',
	WordWdSelectionTypeSelectionNormal = '\002a\000\002',
	WordWdSelectionTypeSelectionFrame = '\002a\000\003',
	WordWdSelectionTypeSelectionColumn = '\002a\000\004',
	WordWdSelectionTypeSelectionRow = '\002a\000\005',
	WordWdSelectionTypeSelectionBlock = '\002a\000\006',
	WordWdSelectionTypeSelectionInlineShape = '\002a\000\007',
	WordWdSelectionTypeSelectionShape = '\002a\000\010'
};
typedef enum WordWdSelectionType WordWdSelectionType;

enum WordWdCaptionLabelID {
	WordWdCaptionLabelIDCaptionFigure = '\002a\377\377',
	WordWdCaptionLabelIDCaptionTable = '\002a\377\376',
	WordWdCaptionLabelIDCaptionEquation = '\002a\377\375'
};
typedef enum WordWdCaptionLabelID WordWdCaptionLabelID;

enum WordWdReferenceType {
	WordWdReferenceTypeReferenceTypeNumberedItem = '\002c\000\000',
	WordWdReferenceTypeReferenceTypeHeading = '\002c\000\001',
	WordWdReferenceTypeReferenceTypeBookmark = '\002c\000\002',
	WordWdReferenceTypeReferenceTypeFootnote = '\002c\000\003',
	WordWdReferenceTypeReferenceTypeEndnote = '\002c\000\004'
};
typedef enum WordWdReferenceType WordWdReferenceType;

enum WordWdReferenceKind {
	WordWdReferenceKindReferenceContentText = '\002c\377\377',
	WordWdReferenceKindReferenceNumberRelativeContext = '\002c\377\376',
	WordWdReferenceKindReferenceNumberNoContext = '\002c\377\375',
	WordWdReferenceKindReferenceNumberFullContext = '\002c\377\374',
	WordWdReferenceKindReferenceEntireCaption = '\002d\000\002',
	WordWdReferenceKindReferenceOnlyLabelAndNumber = '\002d\000\003',
	WordWdReferenceKindReferenceOnlyCaptionText = '\002d\000\004',
	WordWdReferenceKindReferenceFootnoteNumber = '\002d\000\005',
	WordWdReferenceKindReferenceEndnoteNumber = '\002d\000\006',
	WordWdReferenceKindReferencePageNumber = '\002d\000\007',
	WordWdReferenceKindReferencePosition = '\002d\000\017',
	WordWdReferenceKindReferenceFootnoteNumberFormatted = '\002d\000\020',
	WordWdReferenceKindReferenceEndnoteNumberFormatted = '\002d\000\021'
};
typedef enum WordWdReferenceKind WordWdReferenceKind;

enum WordWdIndexFormat {
	WordWdIndexFormatIndexTemplate = '\002e\000\000',
	WordWdIndexFormatIndexClassic = '\002e\000\001',
	WordWdIndexFormatIndexFancy = '\002e\000\002',
	WordWdIndexFormatIndexModern = '\002e\000\003',
	WordWdIndexFormatIndexBulleted = '\002e\000\004',
	WordWdIndexFormatIndexFormal = '\002e\000\005',
	WordWdIndexFormatIndexSimple = '\002e\000\006'
};
typedef enum WordWdIndexFormat WordWdIndexFormat;

enum WordWdIndexType {
	WordWdIndexTypeIndexIndent = '\002f\000\000',
	WordWdIndexTypeIndexRunin = '\002f\000\001'
};
typedef enum WordWdIndexType WordWdIndexType;

enum WordWdRevisionsWrap {
	WordWdRevisionsWrapWrapNever = '\002g\000\000',
	WordWdRevisionsWrapWrapAlways = '\002g\000\001',
	WordWdRevisionsWrapWrapAsk = '\002g\000\002'
};
typedef enum WordWdRevisionsWrap WordWdRevisionsWrap;

enum WordWdRevisionType {
	WordWdRevisionTypeNoRevision = '\002h\000\000',
	WordWdRevisionTypeRevisionInsert = '\002h\000\001',
	WordWdRevisionTypeRevisionDelete = '\002h\000\002',
	WordWdRevisionTypeRevisionProperty = '\002h\000\003',
	WordWdRevisionTypeRevisionParagraphNumber = '\002h\000\004',
	WordWdRevisionTypeRevisionDisplayField = '\002h\000\005',
	WordWdRevisionTypeRevisionReconcile = '\002h\000\006',
	WordWdRevisionTypeRevisionConflict = '\002h\000\007',
	WordWdRevisionTypeRevisionStyle = '\002h\000\010',
	WordWdRevisionTypeRevisionReplace = '\002h\000\011',
	WordWdRevisionTypeRevisionParagraphProperty = '\002h\000\012',
	WordWdRevisionTypeRevisionTableProperty = '\002h\000\013',
	WordWdRevisionTypeRevisionSectionProperty = '\002h\000\014',
	WordWdRevisionTypeRevisionStyleDefinition = '\002h\000\015',
	WordWdRevisionTypeRevisionMoveFrom = '\002h\000\016',
	WordWdRevisionTypeRevisionMoveTo = '\002h\000\017',
	WordWdRevisionTypeRevisionCellInsertion = '\002h\000\020',
	WordWdRevisionTypeRevisionCellDeletion = '\002h\000\021',
	WordWdRevisionTypeRevisionCellMerge = '\002h\000\022',
	WordWdRevisionTypeRevisionCellSplit = '\002h\000\023',
	WordWdRevisionTypeRevisionConflictInsert = '\002h\000\024',
	WordWdRevisionTypeRevisionConflictDelete = '\002h\000\025'
};
typedef enum WordWdRevisionType WordWdRevisionType;

enum WordWdRoutingSlipDelivery {
	WordWdRoutingSlipDeliveryOneAfterAnother = '\002i\000\000',
	WordWdRoutingSlipDeliveryAllAtOnce = '\002i\000\001'
};
typedef enum WordWdRoutingSlipDelivery WordWdRoutingSlipDelivery;

enum WordWdRoutingSlipStatus {
	WordWdRoutingSlipStatusNotYetRouted = '\002j\000\000',
	WordWdRoutingSlipStatusRouteInProgress = '\002j\000\001',
	WordWdRoutingSlipStatusRouteComplete = '\002j\000\002'
};
typedef enum WordWdRoutingSlipStatus WordWdRoutingSlipStatus;

enum WordWdSectionStart {
	WordWdSectionStartSectionContinuous = '\002k\000\000',
	WordWdSectionStartSectionNewColumn = '\002k\000\001',
	WordWdSectionStartSectionNewPage = '\002k\000\002',
	WordWdSectionStartSectionEvenPage = '\002k\000\003',
	WordWdSectionStartSectionOddPage = '\002k\000\004'
};
typedef enum WordWdSectionStart WordWdSectionStart;

enum WordWdSaveOptions {
	WordWdSaveOptionsDoNotSaveChanges = '\002l\000\000',
	WordWdSaveOptionsSaveChanges = '\002k\377\377',
	WordWdSaveOptionsPromptToSaveChanges = '\002k\377\376'
};
typedef enum WordWdSaveOptions WordWdSaveOptions;

enum WordWdDocumentKind {
	WordWdDocumentKindDocumentNotSpecified = '\002m\000\000',
	WordWdDocumentKindDocumentLetter = '\002m\000\001',
	WordWdDocumentKindDocumentEmail = '\002m\000\002'
};
typedef enum WordWdDocumentKind WordWdDocumentKind;

enum WordWdDocumentType {
	WordWdDocumentTypeTypeDocument = '\002n\000\000',
	WordWdDocumentTypeTypeTemplate = '\002n\000\001',
	WordWdDocumentTypeTypeFrameset = '\002n\000\002'
};
typedef enum WordWdDocumentType WordWdDocumentType;

enum WordWdOriginalFormat {
	WordWdOriginalFormatWordDocument = '\002o\000\000',
	WordWdOriginalFormatOriginalDocumentFormat = '\002o\000\001',
	WordWdOriginalFormatPromptUser = '\002o\000\002'
};
typedef enum WordWdOriginalFormat WordWdOriginalFormat;

enum WordWdRelocate {
	WordWdRelocateRelocateUp = '\002p\000\000',
	WordWdRelocateRelocateDown = '\002p\000\001'
};
typedef enum WordWdRelocate WordWdRelocate;

enum WordWdInsertedTextMark {
	WordWdInsertedTextMarkInsertedTextMarkNone = '\002q\000\000',
	WordWdInsertedTextMarkInsertedTextMarkBold = '\002q\000\001',
	WordWdInsertedTextMarkInsertedTextMarkItalic = '\002q\000\002',
	WordWdInsertedTextMarkInsertedTextMarkUnderline = '\002q\000\003',
	WordWdInsertedTextMarkInsertedTextMarkDoubleUnderline = '\002q\000\004',
	WordWdInsertedTextMarkInsertedTextMarkColorOnly = '\002q\000\005',
	WordWdInsertedTextMarkInsertedTextMarkStrikeThrough = '\002q\000\006',
	WordWdInsertedTextMarkInsertedTextMarkDoubleStrikeThrough = '\002q\000\007'
};
typedef enum WordWdInsertedTextMark WordWdInsertedTextMark;

enum WordWdRevisedLinesMark {
	WordWdRevisedLinesMarkRevisedLinesMarkNone = '\002r\000\000',
	WordWdRevisedLinesMarkRevisedLinesMarkLeftBorder = '\002r\000\001',
	WordWdRevisedLinesMarkRevisedLinesMarkRightBorder = '\002r\000\002',
	WordWdRevisedLinesMarkRevisedLinesMarkOutsideBorder = '\002r\000\003'
};
typedef enum WordWdRevisedLinesMark WordWdRevisedLinesMark;

enum WordWdDeletedTextMark {
	WordWdDeletedTextMarkDeletedTextMarkHidden = '\002s\000\000',
	WordWdDeletedTextMarkDeletedTextMarkStrikeThrough = '\002s\000\001',
	WordWdDeletedTextMarkDeletedTextMarkCaret = '\002s\000\002',
	WordWdDeletedTextMarkDeletedTextMarkPound = '\002s\000\003',
	WordWdDeletedTextMarkDeletedTextMarkNone = '\002s\000\004',
	WordWdDeletedTextMarkDeletedTextMarkBold = '\002s\000\005',
	WordWdDeletedTextMarkDeletedTextMarkItalic = '\002s\000\006',
	WordWdDeletedTextMarkDeletedTextMarkUnderline = '\002s\000\007',
	WordWdDeletedTextMarkDeletedTextMarkDoubleUnderline = '\002s\000\010',
	WordWdDeletedTextMarkDeletedTextMarkColorOnly = '\002s\000\011',
	WordWdDeletedTextMarkDeletedTextMarkDoubleStrikeThrough = '\002s\000\012'
};
typedef enum WordWdDeletedTextMark WordWdDeletedTextMark;

enum WordWdRevisedPropertiesMark {
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkNone = '\002t\000\000',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkBold = '\002t\000\001',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkItalic = '\002t\000\002',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkUnderline = '\002t\000\003',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkDoubleUnderline = '\002t\000\004',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkColorOnly = '\002t\000\005',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkStrikeThrough = '\002t\000\006',
	WordWdRevisedPropertiesMarkRevisedPropertiesMarkDoubleStrikeThrough = '\002t\000\007'
};
typedef enum WordWdRevisedPropertiesMark WordWdRevisedPropertiesMark;

enum WordWdMoveToTextMark {
	WordWdMoveToTextMarkMoveToTextMarkNone = '\002\314\000\000',
	WordWdMoveToTextMarkMoveToTextMarkBold = '\002\314\000\001',
	WordWdMoveToTextMarkMoveToTextMarkItalic = '\002\314\000\002',
	WordWdMoveToTextMarkMoveToTextMarkUnderline = '\002\314\000\003',
	WordWdMoveToTextMarkMoveToTextMarkDoubleUnderline = '\002\314\000\004',
	WordWdMoveToTextMarkMoveToTextMarkColorOnly = '\002\314\000\005',
	WordWdMoveToTextMarkMoveToTextMarkStrikeThrough = '\002\314\000\006',
	WordWdMoveToTextMarkMoveToTextMarkDoubleStrikeThrough = '\002\314\000\007'
};
typedef enum WordWdMoveToTextMark WordWdMoveToTextMark;

enum WordWdOMathFunctionType {
	WordWdOMathFunctionTypeMathAccentType = '\002\317\000\001',
	WordWdOMathFunctionTypeMathFunctionBarType = '\002\317\000\002',
	WordWdOMathFunctionTypeMathBoxType = '\002\317\000\003',
	WordWdOMathFunctionTypeMathBoxedFormulaType = '\002\317\000\004',
	WordWdOMathFunctionTypeMathDelimitersType = '\002\317\000\005',
	WordWdOMathFunctionTypeMathEquationArrayType = '\002\317\000\006',
	WordWdOMathFunctionTypeMathFractionType = '\002\317\000\007',
	WordWdOMathFunctionTypeMathFunctionApplyType = '\002\317\000\010',
	WordWdOMathFunctionTypeMathStretchStackType = '\002\317\000\011',
	WordWdOMathFunctionTypeMathLowerLimitType = '\002\317\000\012',
	WordWdOMathFunctionTypeMathUpperLimitType = '\002\317\000\013',
	WordWdOMathFunctionTypeMathMatrixType = '\002\317\000\014',
	WordWdOMathFunctionTypeMathNaryOperatorType = '\002\317\000\015',
	WordWdOMathFunctionTypeMathPhantomType = '\002\317\000\016',
	WordWdOMathFunctionTypeMathLeftSubSuperscriptType = '\002\317\000\017',
	WordWdOMathFunctionTypeMathRadicalType = '\002\317\000\020',
	WordWdOMathFunctionTypeMathSubscriptType = '\002\317\000\021',
	WordWdOMathFunctionTypeMathSubSuperscriptType = '\002\317\000\022',
	WordWdOMathFunctionTypeMathSuperscriptType = '\002\317\000\023',
	WordWdOMathFunctionTypeMathTextType = '\002\317\000\024',
	WordWdOMathFunctionTypeMathNormalTextType = '\002\317\000\025',
	WordWdOMathFunctionTypeMathLiteralTextType = '\002\317\000\026'
};
typedef enum WordWdOMathFunctionType WordWdOMathFunctionType;

enum WordWdOMathHorizAlignType {
	WordWdOMathHorizAlignTypeEquationHorizontalAlignCenter = '\002\320\000\000',
	WordWdOMathHorizAlignTypeEquationHorizontalAlignLeft = '\002\320\000\001',
	WordWdOMathHorizAlignTypeEquationHorizontalAlignRight = '\002\320\000\002'
};
typedef enum WordWdOMathHorizAlignType WordWdOMathHorizAlignType;

enum WordWdOMathVertAlignType {
	WordWdOMathVertAlignTypeEquationVerticalAlignCenter = '\002\321\000\000',
	WordWdOMathVertAlignTypeEquationVerticalAlignTop = '\002\321\000\001',
	WordWdOMathVertAlignTypeEquationVerticalAlignBottom = '\002\321\000\002'
};
typedef enum WordWdOMathVertAlignType WordWdOMathVertAlignType;

enum WordWdOMathFracType {
	WordWdOMathFracTypeMathFractionTypeBar = '\002\322\000\000',
	WordWdOMathFracTypeMathFractionTypeStack = '\002\322\000\001',
	WordWdOMathFracTypeMathFractionTypeSlashed = '\002\322\000\002',
	WordWdOMathFracTypeMathFractionTypeLinear = '\002\322\000\003'
};
typedef enum WordWdOMathFracType WordWdOMathFracType;

enum WordWdOMathSpacingRule {
	WordWdOMathSpacingRuleEquationSpacingSingle = '\002\323\000\000',
	WordWdOMathSpacingRuleEquationSpacingOneAndAHalf = '\002\323\000\001',
	WordWdOMathSpacingRuleEquationSpacingDouble = '\002\323\000\002',
	WordWdOMathSpacingRuleEquationSpacingExactly = '\002\323\000\003',
	WordWdOMathSpacingRuleEquationSpacingMultiple = '\002\323\000\004'
};
typedef enum WordWdOMathSpacingRule WordWdOMathSpacingRule;

enum WordWdOMathType {
	WordWdOMathTypeEquationDisplayProfessional = '\002\324\000\000' /* Specifies an equation is displayed in professional/built up format. */,
	WordWdOMathTypeEquationDisplayInline = '\002\324\000\001' /* Specifies an equation is displayed in linear/built down style. */
};
typedef enum WordWdOMathType WordWdOMathType;

enum WordWdOMathShapeType {
	WordWdOMathShapeTypeMathDelimitersCenterVertically = '\002\325\000\000',
	WordWdOMathShapeTypeMathDelimitersMatchContentSize = '\002\325\000\001'
};
typedef enum WordWdOMathShapeType WordWdOMathShapeType;

enum WordWdOMathJc {
	WordWdOMathJcEquationJustificationCenterGroup = '\002\326\000\001',
	WordWdOMathJcEquationJustificationCenter = '\002\326\000\002',
	WordWdOMathJcEquationJustificationLeft = '\002\326\000\003',
	WordWdOMathJcEquationJustificationRight = '\002\326\000\004',
	WordWdOMathJcEquationJustificationInline = '\002\326\000\007'
};
typedef enum WordWdOMathJc WordWdOMathJc;

enum WordWdOMathBreakBin {
	WordWdOMathBreakBinMathBinaryOperatorBreakBefore = '\002\327\000\000',
	WordWdOMathBreakBinMathBinaryOperatorBreakAfter = '\002\327\000\001',
	WordWdOMathBreakBinMathBinaryOperatorRepeat = '\002\327\000\002'
};
typedef enum WordWdOMathBreakBin WordWdOMathBreakBin;

enum WordWdOMathBreakSub {
	WordWdOMathBreakSubMinusMinus = '\002\330\000\000',
	WordWdOMathBreakSubPlusMinus = '\002\330\000\001',
	WordWdOMathBreakSubMinusPlus = '\002\330\000\002'
};
typedef enum WordWdOMathBreakSub WordWdOMathBreakSub;

enum WordWdBuildingBlockTypes {
	WordWdBuildingBlockTypesBuildingBlockEquations = '\000\000\000\003'
};
typedef enum WordWdBuildingBlockTypes WordWdBuildingBlockTypes;

enum WordWdDocPartInsertOptions {
	WordWdDocPartInsertOptionsInlineBuildingBlock = '\000\000\000\000',
	WordWdDocPartInsertOptionsPageLevelBuildingBlock = '\000\000\000\001',
	WordWdDocPartInsertOptionsParagraphLevelBuildingBlock = '\000\000\000\002'
};
typedef enum WordWdDocPartInsertOptions WordWdDocPartInsertOptions;

enum WordWdMoveFromTextMark {
	WordWdMoveFromTextMarkMoveFromTextMarkHidden = '\002\315\000\000',
	WordWdMoveFromTextMarkMoveFromTextMarkDoubleStrikeThrough = '\002\315\000\001',
	WordWdMoveFromTextMarkMoveFromTextMarkStrikeThrough = '\002\315\000\002',
	WordWdMoveFromTextMarkMoveFromTextMarkCaret = '\002\315\000\003',
	WordWdMoveFromTextMarkMoveFromTextMarkPound = '\002\315\000\004',
	WordWdMoveFromTextMarkMoveFromTextMarkNone = '\002\315\000\005',
	WordWdMoveFromTextMarkMoveFromTextMarkBold = '\002\315\000\006',
	WordWdMoveFromTextMarkMoveFromTextMarkItalic = '\002\315\000\007',
	WordWdMoveFromTextMarkMoveFromTextMarkUnderline = '\002\315\000\010',
	WordWdMoveFromTextMarkMoveFromTextMarkDoubleUnderline = '\002\315\000\011',
	WordWdMoveFromTextMarkMoveFromTextMarkColorOnly = '\002\315\000\012'
};
typedef enum WordWdMoveFromTextMark WordWdMoveFromTextMark;

enum WordWdCellColor {
	WordWdCellColorCellColorByAuthor = '\002\315\377\377',
	WordWdCellColorCellColorNoHighlight = '\002\316\000\000',
	WordWdCellColorCellColorPink = '\002\316\000\001',
	WordWdCellColorCellColorLightBlue = '\002\316\000\002',
	WordWdCellColorCellColorLightYellow = '\002\316\000\003',
	WordWdCellColorCellColorLightPurple = '\002\316\000\004',
	WordWdCellColorCellColorLightOrange = '\002\316\000\005',
	WordWdCellColorCellColorLightGreen = '\002\316\000\006',
	WordWdCellColorCellColorLightGray = '\002\316\000\007'
};
typedef enum WordWdCellColor WordWdCellColor;

enum WordWdFieldShading {
	WordWdFieldShadingFieldShadingNever = '\002u\000\000',
	WordWdFieldShadingFieldShadingAlways = '\002u\000\001',
	WordWdFieldShadingFieldShadingWhenSelected = '\002u\000\002'
};
typedef enum WordWdFieldShading WordWdFieldShading;

enum WordWdDefaultFilePath {
	WordWdDefaultFilePathDocumentsPath = '\002v\000\000',
	WordWdDefaultFilePathPicturesPath = '\002v\000\001',
	WordWdDefaultFilePathUserTemplatesPath = '\002v\000\002',
	WordWdDefaultFilePathWorkgroupTemplatesPath = '\002v\000\003',
	WordWdDefaultFilePathUserOptionsPath = '\002v\000\004',
	WordWdDefaultFilePathAutoRecoverPath = '\002v\000\005',
	WordWdDefaultFilePathToolsPath = '\002v\000\006',
	WordWdDefaultFilePathTutorialPath = '\002v\000\007',
	WordWdDefaultFilePathStartupPath = '\002v\000\010',
	WordWdDefaultFilePathProgramPath = '\002v\000\011',
	WordWdDefaultFilePathGraphicsFiltersPath = '\002v\000\012',
	WordWdDefaultFilePathTextConvertersPath = '\002v\000\013',
	WordWdDefaultFilePathProofingToolsPath = '\002v\000\014',
	WordWdDefaultFilePathTempFilePath = '\002v\000\015',
	WordWdDefaultFilePathCurrentFolderPath = '\002v\000\016',
	WordWdDefaultFilePathStyleGalleryPath = '\002v\000\017',
	WordWdDefaultFilePathBorderArtPath = '\002v\000\023'
};
typedef enum WordWdDefaultFilePath WordWdDefaultFilePath;

enum WordWdCompatibility {
	WordWdCompatibilityNoTabHangingIndent = '\002w\000\001',
	WordWdCompatibilityNoSpaceForRaisedOrLoweredCharacters = '\002w\000\002',
	WordWdCompatibilityPrintColorsBlack = '\002w\000\003',
	WordWdCompatibilityWrapTrailSpaces = '\002w\000\004',
	WordWdCompatibilityNoColumnBalance = '\002w\000\005',
	WordWdCompatibilityConvertDataMergeEscapes = '\002w\000\006',
	WordWdCompatibilitySuppressSpaceBeforeAfterPageBreak = '\002w\000\007',
	WordWdCompatibilitySuppressTopSpacing = '\002w\000\010',
	WordWdCompatibilityOriginalWordTableRules = '\002w\000\011',
	WordWdCompatibilityTransparentMetafiles = '\002w\000\012',
	WordWdCompatibilityShowBreaksInFrames = '\002w\000\013',
	WordWdCompatibilitySwapBordersFacingPages = '\002w\000\014',
	WordWdCompatibilityLeaveBackslashAlone = '\002w\000\015',
	WordWdCompatibilityExpandShiftReturn = '\002w\000\016',
	WordWdCompatibilityDoNotUnderlineTrailingSpaces = '\002w\000\017',
	WordWdCompatibilityDoNotBalanceSBCSAndDBCSCharacters = '\002w\000\020',
	WordWdCompatibilitySuppressTopSpacingMacWord5 = '\002w\000\021',
	WordWdCompatibilitySpacingInWholePoints = '\002w\000\022',
	WordWdCompatibilityPrintBodyTextBeforeHeader = '\002w\000\023',
	WordWdCompatibilityNoExtraSpaceBetweenRowsOfText = '\002w\000\024',
	WordWdCompatibilityNoSpaceForUnderlines = '\002w\000\025',
	WordWdCompatibilityUseLargerSmallCaps = '\002w\000\026',
	WordWdCompatibilityNoExtraLineSpacing = '\002w\000\027',
	WordWdCompatibilityTruncateFontHeight = '\002w\000\030',
	WordWdCompatibilitySubstituteFontBySize = '\002w\000\031',
	WordWdCompatibilityUsePrinterMetrics = '\002w\000\032',
	WordWdCompatibilityWord6BorderRules = '\002w\000\033',
	WordWdCompatibilityExactOnTop = '\002w\000\034',
	WordWdCompatibilitySuppressBottomSpacing = '\002w\000\035',
	WordWdCompatibilityWordPerfectSpaceWidth = '\002w\000\036',
	WordWdCompatibilityWordPerfectJustification = '\002w\000\037',
	WordWdCompatibilityWord6LineWrap = '\002w\000 ',
	WordWdCompatibilityWord96ShapeLayout = '\002w\000!',
	WordWdCompatibilityWord98FootnoteLayout = '\002w\000\"',
	WordWdCompatibilityDoNotUseHtmlParagraphAutoSpacing = '\002w\000#',
	WordWdCompatibilityDoNotAdjustLineHeightInTable = '\002w\000$',
	WordWdCompatibilityForgetLastTabAlignment = '\002w\000%',
	WordWdCompatibilityWord95AutoSpace = '\002w\000&',
	WordWdCompatibilityAlignTablesRowByRow = '\002w\000\'',
	WordWdCompatibilityLayoutRawTableWidth = '\002w\000(',
	WordWdCompatibilityLayoutTableRowsApart = '\002w\000)',
	WordWdCompatibilityUseWord97LineBreakingRules = '\002w\000*',
	WordWdCompatibilityDoNotBreakWrappedTables = '\002w\000+',
	WordWdCompatibilityDoNotSnapTextToGridInTableWithObjects = '\002w\000,',
	WordWdCompatibilitySelectFieldWithFirstOrLastCharacter = '\002w\000-',
	WordWdCompatibilityApplyBreakingRules = '\002w\000.',
	WordWdCompatibilityDoNotWrapTextWithPunctuation = '\002w\000/',
	WordWdCompatibilityDoNotUseAsianBreakRulesInGrid = '\002w\0000',
	WordWdCompatibilityUseWord2002TableStyleRules = '\002w\0001',
	WordWdCompatibilityGrowAutofit = '\002w\0002',
	WordWdCompatibilityUseNormalStyleForList = '\002w\0003',
	WordWdCompatibilityDoNotUseIndentAsNumberingTabStop = '\002w\0004',
	WordWdCompatibilityLineBreak = '\002w\0005',
	WordWdCompatibilityAllowSpaceOfSameStyleInTable = '\002w\0006',
	WordWdCompatibilityIndentRules = '\002w\0007',
	WordWdCompatibilityDontAutofitConstrainedTables = '\002w\0008',
	WordWdCompatibilityAutofitLike = '\002w\0009',
	WordWdCompatibilityUnderlineTabInNumList = '\002w\000:',
	WordWdCompatibilityHangulWidthLike = '\002w\000;',
	WordWdCompatibilitySplitPgBreakAndParaMark = '\002w\000<',
	WordWdCompatibilityDontVertAlignCellWithShape = '\002w\000=',
	WordWdCompatibilityDontBreakConstrainedForcedTables = '\002w\000>',
	WordWdCompatibilityDontVertAlignInTextbox = '\002w\000\?',
	WordWdCompatibilityWordKerningPairs = '\002w\000@',
	WordWdCompatibilityCachedColBalance = '\002w\000A',
	WordWdCompatibilityDisableKerning = '\002w\000B',
	WordWdCompatibilityFlipMirrorIndents = '\002w\000C',
	WordWdCompatibilityDontOverrideTableStyleFontSzAndJustification = '\002w\000D',
	WordWdCompatibilityUseWordTableStyleRules = '\002w\000E',
	WordWdCompatibilityDelayableFloatingTable = '\002w\000F',
	WordWdCompatibilityAllowHyphenationAtTrackBottom = '\002w\000G',
	WordWdCompatibilityUseWordTrackBottomHyphenation = '\002w\000H',
	WordWdCompatibilityUsePre2018WordMacFontLayout = '\002w\000I'
};
typedef enum WordWdCompatibility WordWdCompatibility;

enum WordWdCompatibilityVersion {
	WordWdCompatibilityVersionDefaultCompatibilitySettings = '\002\312\000\000' /* Microsoft Word 2007-2008 */,
	WordWdCompatibilityVersionWord20002004AndX = '\002\312\000\014' /* Microsoft Word 2000-2004 and X */,
	WordWdCompatibilityVersionWord97And98 = '\002\312\000\001' /* Microsoft Word 97-98 */,
	WordWdCompatibilityVersionWord6And95 = '\002\312\000\002' /* Microsoft Word 6.0/95 */,
	WordWdCompatibilityVersionWinWord1 = '\002\312\000\003' /* Word for Windows 1.0 */,
	WordWdCompatibilityVersionWinWord2 = '\002\312\000\004' /* Word for Windows 2.0 */,
	WordWdCompatibilityVersionMacWord5 = '\002\312\000\005' /* Word for the Macintosh 5.x */,
	WordWdCompatibilityVersionDosWord = '\002\312\000\006' /* Word for MS-DOS */,
	WordWdCompatibilityVersionWordperfect5 = '\002\312\000\007' /* WordPerfect 5.x */,
	WordWdCompatibilityVersionWinWordperfect6 = '\002\312\000\010' /* WordPerfect 6.x for Windows */,
	WordWdCompatibilityVersionDosWordperfect6 = '\002\312\000\011' /* WordPerfect 6.0 for DOS */,
	WordWdCompatibilityVersionAsianWord97And98 = '\002\312\000\012' /* Microsoft Word Asian Versions 97-98 */,
	WordWdCompatibilityVersionUsWord6And95 = '\002\312\000\013' /* US Microsoft Word for Windows 6.0/95 */,
	WordWdCompatibilityVersionCustomCompatibilitySettings = '\002\312\000\015' /* Custom settings */
};
typedef enum WordWdCompatibilityVersion WordWdCompatibilityVersion;

enum WordWdPaperSize {
	WordWdPaperSizePaperTenXFourteen = '\002x\000\000',
	WordWdPaperSizePaperElevenXSeventeen = '\002x\000\001',
	WordWdPaperSizePaperLetter = '\002x\000\002',
	WordWdPaperSizePaperLetterSmall = '\002x\000\003',
	WordWdPaperSizePaperLegal = '\002x\000\004',
	WordWdPaperSizePaperExecutive = '\002x\000\005',
	WordWdPaperSizePaperA3 = '\002x\000\006',
	WordWdPaperSizePaperA4 = '\002x\000\007',
	WordWdPaperSizePaperA4Small = '\002x\000\010',
	WordWdPaperSizePaperA5 = '\002x\000\011',
	WordWdPaperSizePaperB4 = '\002x\000\012',
	WordWdPaperSizePaperB5 = '\002x\000\013',
	WordWdPaperSizePaperCsheet = '\002x\000\014',
	WordWdPaperSizePaperDsheet = '\002x\000\015',
	WordWdPaperSizePaperEsheet = '\002x\000\016',
	WordWdPaperSizePaperFanfoldLegalGerman = '\002x\000\017',
	WordWdPaperSizePaperFanfoldStdGerman = '\002x\000\020',
	WordWdPaperSizePaperFanfoldUs = '\002x\000\021',
	WordWdPaperSizePaperFolio = '\002x\000\022',
	WordWdPaperSizePaperLedger = '\002x\000\023',
	WordWdPaperSizePaperNote = '\002x\000\024',
	WordWdPaperSizePaperQuarto = '\002x\000\025',
	WordWdPaperSizePaperStatement = '\002x\000\026',
	WordWdPaperSizePaperTabloid = '\002x\000\027',
	WordWdPaperSizePaperEnvelope9 = '\002x\000\030',
	WordWdPaperSizePaperEnvelope10 = '\002x\000\031',
	WordWdPaperSizePaperEnvelope11 = '\002x\000\032',
	WordWdPaperSizePaperEnvelope12 = '\002x\000\033',
	WordWdPaperSizePaperEnvelope14 = '\002x\000\034',
	WordWdPaperSizePaperEnvelopeB4 = '\002x\000\035',
	WordWdPaperSizePaperEnvelopeB5 = '\002x\000\036',
	WordWdPaperSizePaperEnvelopeB6 = '\002x\000\037',
	WordWdPaperSizePaperEnvelopeC3 = '\002x\000 ',
	WordWdPaperSizePaperEnvelopeC4 = '\002x\000!',
	WordWdPaperSizePaperEnvelopeC5 = '\002x\000\"',
	WordWdPaperSizePaperEnvelopeC6 = '\002x\000#',
	WordWdPaperSizePaperEnvelopeC65 = '\002x\000$',
	WordWdPaperSizePaperEnvelopeDl = '\002x\000%',
	WordWdPaperSizePaperEnvelopeItaly = '\002x\000&',
	WordWdPaperSizePaperEnvelopeMonarch = '\002x\000\'',
	WordWdPaperSizePaperEnvelopePersonal = '\002x\000(',
	WordWdPaperSizePaperCustom = '\002x\000)'
};
typedef enum WordWdPaperSize WordWdPaperSize;

enum WordWdCustomLabelPageSize {
	WordWdCustomLabelPageSizeCustomLabelLetter = '\002y\000\000',
	WordWdCustomLabelPageSizeCustomLabelLetterLandscape = '\002y\000\001',
	WordWdCustomLabelPageSizeCustomLabelA4 = '\002y\000\002',
	WordWdCustomLabelPageSizeCustomLabelA4Landscape = '\002y\000\003',
	WordWdCustomLabelPageSizeCustomLabelA5 = '\002y\000\004',
	WordWdCustomLabelPageSizeCustomLabelA5Landscape = '\002y\000\005',
	WordWdCustomLabelPageSizeCustomLabelB5 = '\002y\000\006',
	WordWdCustomLabelPageSizeCustomLabelMini = '\002y\000\007',
	WordWdCustomLabelPageSizeCustomLabelFanfold = '\002y\000\010',
	WordWdCustomLabelPageSizeCustomLabelVertHalfSheet = '\002y\000\011',
	WordWdCustomLabelPageSizeCustomLabelVertHalfSheetLs = '\002y\000\012',
	WordWdCustomLabelPageSizeCustomLabelHigaki = '\002y\000\013',
	WordWdCustomLabelPageSizeCustomLabelHigakiLs = '\002y\000\014',
	WordWdCustomLabelPageSizeCustomLabelB4Jis = '\002y\000\015'
};
typedef enum WordWdCustomLabelPageSize WordWdCustomLabelPageSize;

enum WordWdProtectionType {
	WordWdProtectionTypeNoDocumentProtection = '\002y\377\377',
	WordWdProtectionTypeAllowOnlyRevisions = '\002z\000\000',
	WordWdProtectionTypeAllowOnlyComments = '\002z\000\001',
	WordWdProtectionTypeAllowOnlyFormFields = '\002z\000\002',
	WordWdProtectionTypeAllowOnlyReading = '\002z\000\003'
};
typedef enum WordWdProtectionType WordWdProtectionType;

enum WordWdPartOfSpeech {
	WordWdPartOfSpeechAdjective = '\002{\000\000',
	WordWdPartOfSpeechNoun = '\002{\000\001',
	WordWdPartOfSpeechAdverb = '\002{\000\002',
	WordWdPartOfSpeechVerb = '\002{\000\003',
	WordWdPartOfSpeechPronoun = '\002{\000\004',
	WordWdPartOfSpeechConjunction = '\002{\000\005',
	WordWdPartOfSpeechPreposition = '\002{\000\006',
	WordWdPartOfSpeechInterjection = '\002{\000\007',
	WordWdPartOfSpeechIdiom = '\002{\000\010',
	WordWdPartOfSpeechOther = '\002{\000\011'
};
typedef enum WordWdPartOfSpeech WordWdPartOfSpeech;

enum WordWdRelativeHorizontalPosition {
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionMargin = '\002|\000\000',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionPage = '\002|\000\001',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionColumn = '\002|\000\002',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionCharacter = '\002|\000\003',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionLeftMargin = '\002|\000\004',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionRightMargin = '\002|\000\005',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionInnerMargin = '\002|\000\006',
	WordWdRelativeHorizontalPositionRelativeHorizontalPositionOuterMargin = '\002|\000\007'
};
typedef enum WordWdRelativeHorizontalPosition WordWdRelativeHorizontalPosition;

enum WordWdRelativeVerticalPosition {
	WordWdRelativeVerticalPositionRelativeVerticalPositionMargin = '\002}\000\000',
	WordWdRelativeVerticalPositionRelativeVerticalPositionPage = '\002}\000\001',
	WordWdRelativeVerticalPositionRelativeVerticalPositionParagraph = '\002}\000\002',
	WordWdRelativeVerticalPositionRelativeVerticalPositionLine = '\002}\000\003',
	WordWdRelativeVerticalPositionRelativeVerticalPositionTopMargin = '\002}\000\004',
	WordWdRelativeVerticalPositionRelativeVerticalPositionBottomMargin = '\002}\000\005',
	WordWdRelativeVerticalPositionRelativeVerticalPositionInnerMargin = '\002}\000\006',
	WordWdRelativeVerticalPositionRelativeVerticalPositionOuterMargin = '\002}\000\007'
};
typedef enum WordWdRelativeVerticalPosition WordWdRelativeVerticalPosition;

enum WordWdHelpType {
	WordWdHelpTypeHelp = '\002~\000\000',
	WordWdHelpTypeHelpAbout = '\002~\000\001',
	WordWdHelpTypeHelpContents = '\002~\000\003',
	WordWdHelpTypeHelpIndex = '\002~\000\005',
	WordWdHelpTypeHelpPsshelp = '\002~\000\007',
	WordWdHelpTypeHelpSearch = '\002~\000\011'
};
typedef enum WordWdHelpType WordWdHelpType;

enum WordWdKeyCategory {
	WordWdKeyCategoryKeyCategoryNil = '\002~\377\377',
	WordWdKeyCategoryKeyCategoryDisable = '\002\177\000\000',
	WordWdKeyCategoryKeyCategoryCommand = '\002\177\000\001',
	WordWdKeyCategoryKeyCategoryMacro = '\002\177\000\002',
	WordWdKeyCategoryKeyCategoryFont = '\002\177\000\003',
	WordWdKeyCategoryKeyCategoryAutoText = '\002\177\000\004',
	WordWdKeyCategoryKeyCategoryStyle = '\002\177\000\005',
	WordWdKeyCategoryKeyCategorySymbol = '\002\177\000\006',
	WordWdKeyCategoryKeyCategoryPrefix = '\002\177\000\007'
};
typedef enum WordWdKeyCategory WordWdKeyCategory;

enum WordWdKey {
	WordWdKeyNo_key = '\002\200\000\377',
	WordWdKeyShift_key = '\002\200\002\000',
	WordWdKeyControl_key = '\002\200\020\000',
	WordWdKeyCommand_key = '\002\200\001\000',
	WordWdKeyOption_key = '\002\200\010\000',
	WordWdKeyKeyAlt = '\002\200\010\000',
	WordWdKeyA_key = '\002\200\000A',
	WordWdKeyB_key = '\002\200\000B',
	WordWdKeyC_key = '\002\200\000C',
	WordWdKeyD_key = '\002\200\000D',
	WordWdKeyE_key = '\002\200\000E',
	WordWdKeyF_key = '\002\200\000F',
	WordWdKeyG_key = '\002\200\000G',
	WordWdKeyH_key = '\002\200\000H',
	WordWdKeyI_Key = '\002\200\000I',
	WordWdKeyJ_key = '\002\200\000J',
	WordWdKeyK_key = '\002\200\000K',
	WordWdKeyL_key = '\002\200\000L',
	WordWdKeyM_key = '\002\200\000M',
	WordWdKeyN_key = '\002\200\000N',
	WordWdKeyO_key = '\002\200\000O',
	WordWdKeyP_key = '\002\200\000P',
	WordWdKeyQ_key = '\002\200\000Q',
	WordWdKeyR_key = '\002\200\000R',
	WordWdKeyS_key = '\002\200\000S',
	WordWdKeyT_key = '\002\200\000T',
	WordWdKeyU_kwy = '\002\200\000U',
	WordWdKeyV_key = '\002\200\000V',
	WordWdKeyW_key = '\002\200\000W',
	WordWdKeyX_key = '\002\200\000X',
	WordWdKeyY_key = '\002\200\000Y',
	WordWdKeyZ_key = '\002\200\000Z',
	WordWdKeyKey_number_0 = '\002\200\0000',
	WordWdKeyKey_number_1 = '\002\200\0001',
	WordWdKeyKey_numbe_2 = '\002\200\0002',
	WordWdKeyKey_numbe_3 = '\002\200\0003',
	WordWdKeyKey_number_4 = '\002\200\0004',
	WordWdKeyKey_number_5 = '\002\200\0005',
	WordWdKeyKey_number_6 = '\002\200\0006',
	WordWdKeyKeyNumber_7 = '\002\200\0007',
	WordWdKeyKey_number_8 = '\002\200\0008',
	WordWdKeyKey_number_9 = '\002\200\0009',
	WordWdKeyBackspace_key = '\002\200\000\010',
	WordWdKeyTab_key = '\002\200\000\011',
	WordWdKeyKey_numeric_5_special = '\002\200\000\014',
	WordWdKeyReturn_key = '\002\200\000\015',
	WordWdKeyPause_key = '\002\200\000\023',
	WordWdKeyEsc_key = '\002\200\000\033',
	WordWdKeySpacebar_key = '\002\200\000 ',
	WordWdKeyPage_up_key = '\002\200\000!',
	WordWdKeyPage_down_key = '\002\200\000\"',
	WordWdKeyEnd_key = '\002\200\000#',
	WordWdKeyHome_key = '\002\200\000$',
	WordWdKeyInsert_key = '\002\200\000-',
	WordWdKeyDelete_key = '\002\200\000.',
	WordWdKeyKey_numeric_0 = '\002\200\000`',
	WordWdKeyKey_numeric_1 = '\002\200\000a',
	WordWdKeyKey_numeric_2 = '\002\200\000b',
	WordWdKeyKey_numeric_3 = '\002\200\000c',
	WordWdKeyKey_numeric_4 = '\002\200\000d',
	WordWdKeyKey_numeric_5 = '\002\200\000e',
	WordWdKeyKey_numeric_6 = '\002\200\000f',
	WordWdKeyKey_numeric_7 = '\002\200\000g',
	WordWdKeyKey_numeric_8 = '\002\200\000h',
	WordWdKeyKey_numeric_9 = '\002\200\000i',
	WordWdKeyKey_numeric_multiply = '\002\200\000j',
	WordWdKeyKey_numeric_add = '\002\200\000k',
	WordWdKeyKey_numeric_subtract = '\002\200\000m',
	WordWdKeyKey_numeric_decimal = '\002\200\000n',
	WordWdKeyKey_numeric_divide = '\002\200\000o',
	WordWdKeyF1_key = '\002\200\000p',
	WordWdKeyF2_key = '\002\200\000q',
	WordWdKeyF3_key = '\002\200\000r',
	WordWdKeyF4_key = '\002\200\000s',
	WordWdKeyF5_key = '\002\200\000t',
	WordWdKeyF6_key = '\002\200\000u',
	WordWdKeyF7_key = '\002\200\000v',
	WordWdKeyF8_key = '\002\200\000w',
	WordWdKeyF9_key = '\002\200\000x',
	WordWdKeyF10_key = '\002\200\000y',
	WordWdKeyF11_key = '\002\200\000z',
	WordWdKeyF12_key = '\002\200\000{',
	WordWdKeyF13_key = '\002\200\000|',
	WordWdKeyF14_key = '\002\200\000}',
	WordWdKeyF15_key = '\002\200\000~',
	WordWdKeyF16_key = '\002\200\000\177',
	WordWdKeyScroll_lock_key = '\002\200\000\221',
	WordWdKeySemi_colon_key = '\002\200\000\272',
	WordWdKeyEquals_key = '\002\200\000\273',
	WordWdKeyComma_key = '\002\200\000\274',
	WordWdKeyHyphen_key = '\002\200\000\275',
	WordWdKeyPeriod_key = '\002\200\000\276',
	WordWdKeySlash_key = '\002\200\000\277',
	WordWdKeyBack_single_quote_key = '\002\200\000\300',
	WordWdKeyOpen_square_brace_key = '\002\200\000\333',
	WordWdKeyBack_slash_key = '\002\200\000\334',
	WordWdKeyClose_square_brace_key = '\002\200\000\335',
	WordWdKeySingle_quote_key = '\002\200\000\336'
};
typedef enum WordWdKey WordWdKey;

enum WordWdOLEType {
	WordWdOLETypeOlelink = '\002\201\000\000',
	WordWdOLETypeOleembed = '\002\201\000\001',
	WordWdOLETypeOlecontrol = '\002\201\000\002'
};
typedef enum WordWdOLEType WordWdOLEType;

enum WordWdOLEVerb {
	WordWdOLEVerbOleverbPrimary = '\002\202\000\000',
	WordWdOLEVerbOleverbShow = '\002\201\377\377',
	WordWdOLEVerbOleverbOpen = '\002\201\377\376',
	WordWdOLEVerbOleverbHide = '\002\201\377\375',
	WordWdOLEVerbOleverbUiactivate = '\002\201\377\374',
	WordWdOLEVerbOleverbInPlaceActivate = '\002\201\377\373',
	WordWdOLEVerbOleverbDiscardUndoState = '\002\201\377\372'
};
typedef enum WordWdOLEVerb WordWdOLEVerb;

enum WordWdOLEPlacement {
	WordWdOLEPlacementInLine = '\002\203\000\000',
	WordWdOLEPlacementFloatOverText = '\002\203\000\001'
};
typedef enum WordWdOLEPlacement WordWdOLEPlacement;

enum WordWdEnvelopeOrientation {
	WordWdEnvelopeOrientationLeftPortrait = '\002\204\000\000',
	WordWdEnvelopeOrientationCenterPortrait = '\002\204\000\001',
	WordWdEnvelopeOrientationRightPortrait = '\002\204\000\002',
	WordWdEnvelopeOrientationLeftLandscape = '\002\204\000\003',
	WordWdEnvelopeOrientationCenterLandscape = '\002\204\000\004',
	WordWdEnvelopeOrientationRightLandscape = '\002\204\000\005',
	WordWdEnvelopeOrientationLeftClockwise = '\002\204\000\006',
	WordWdEnvelopeOrientationCenterClockwise = '\002\204\000\007',
	WordWdEnvelopeOrientationRightClockwise = '\002\204\000\010'
};
typedef enum WordWdEnvelopeOrientation WordWdEnvelopeOrientation;

enum WordWdLetterStyle {
	WordWdLetterStyleFullBlock = '\002\205\000\000',
	WordWdLetterStyleModifiedBlock = '\002\205\000\001',
	WordWdLetterStyleSemiBlock = '\002\205\000\002'
};
typedef enum WordWdLetterStyle WordWdLetterStyle;

enum WordWdLetterheadLocation {
	WordWdLetterheadLocationLetterTop = '\002\206\000\000',
	WordWdLetterheadLocationLetterBottom = '\002\206\000\001',
	WordWdLetterheadLocationLetterLeft = '\002\206\000\002',
	WordWdLetterheadLocationLetterRight = '\002\206\000\003'
};
typedef enum WordWdLetterheadLocation WordWdLetterheadLocation;

enum WordWdSalutationType {
	WordWdSalutationTypeSalutationInformal = '\002\207\000\000',
	WordWdSalutationTypeSalutationFormal = '\002\207\000\001',
	WordWdSalutationTypeSalutationBusiness = '\002\207\000\002',
	WordWdSalutationTypeSalutationOther = '\002\207\000\003'
};
typedef enum WordWdSalutationType WordWdSalutationType;

enum WordWdSalutationGender {
	WordWdSalutationGenderGenderFemale = '\002\210\000\000',
	WordWdSalutationGenderGenderMale = '\002\210\000\001',
	WordWdSalutationGenderGenderNeutral = '\002\210\000\002',
	WordWdSalutationGenderGenderUnknown = '\002\210\000\003'
};
typedef enum WordWdSalutationGender WordWdSalutationGender;

enum WordWdMovementType {
	WordWdMovementTypeByMoving = '\002\211\000\000',
	WordWdMovementTypeBySelecting = '\002\211\000\001'
};
typedef enum WordWdMovementType WordWdMovementType;

enum WordWdConstants {
	WordWdConstantsUndefinedConstant = '\000\230\226\177',
	WordWdConstantsToggleConstant = '\000\230\226~',
	WordWdConstantsForwardConstant = '\?\377\377\377',
	WordWdConstantsBackwardConstant = '\300\000\000\001',
	WordWdConstantsAutoPosition = '\000\000\000\000',
	WordWdConstantsFirstConstant = '\000\000\000\001',
	WordWdConstantsCreatorCode = 'MSWD'
};
typedef enum WordWdConstants WordWdConstants;

enum WordWdPasteDataType {
	WordWdPasteDataTypePasteOleobject = '\002\213\000\000',
	WordWdPasteDataTypePasteRtf = '\002\213\000\001',
	WordWdPasteDataTypePasteText = '\002\213\000\002',
	WordWdPasteDataTypePasteMetafilePicture = '\002\213\000\003',
	WordWdPasteDataTypePasteBitmap = '\002\213\000\004',
	WordWdPasteDataTypePasteDeviceIndependentBitmap = '\002\213\000\005',
	WordWdPasteDataTypePasteHyperlink = '\002\213\000\007',
	WordWdPasteDataTypePasteShape = '\002\213\000\010',
	WordWdPasteDataTypePasteEnhancedMetafile = '\002\213\000\011',
	WordWdPasteDataTypePasteHtml = '\002\213\000\012'
};
typedef enum WordWdPasteDataType WordWdPasteDataType;

enum WordWdPrintOutItem {
	WordWdPrintOutItemPrintDocumentContent = '\002\214\000\000',
	WordWdPrintOutItemPrintProperties = '\002\214\000\001',
	WordWdPrintOutItemPrintComments = '\002\214\000\002',
	WordWdPrintOutItemPrintMarkup = '\002\214\000\002',
	WordWdPrintOutItemPrintStyles = '\002\214\000\003',
	WordWdPrintOutItemPrintAutoTextEntries = '\002\214\000\004',
	WordWdPrintOutItemPrintKeyAssignments = '\002\214\000\005',
	WordWdPrintOutItemPrintEnvelope = '\002\214\000\006',
	WordWdPrintOutItemPrintDocumentWithMarkup = '\002\214\000\007'
};
typedef enum WordWdPrintOutItem WordWdPrintOutItem;

enum WordWdPrintOutPages {
	WordWdPrintOutPagesPrintAllPages = '\002\215\000\000',
	WordWdPrintOutPagesPrintOddPagesOnly = '\002\215\000\001',
	WordWdPrintOutPagesPrintEvenPagesOnly = '\002\215\000\002'
};
typedef enum WordWdPrintOutPages WordWdPrintOutPages;

enum WordWdPrintOutRange {
	WordWdPrintOutRangePrintAllDocument = '\002\216\000\000',
	WordWdPrintOutRangePrintSelection = '\002\216\000\001',
	WordWdPrintOutRangePrintCurrentPage = '\002\216\000\002',
	WordWdPrintOutRangePrintFromTo = '\002\216\000\003',
	WordWdPrintOutRangePrintRangeOfPages = '\002\216\000\004'
};
typedef enum WordWdPrintOutRange WordWdPrintOutRange;

enum WordWdDictionaryType {
	WordWdDictionaryTypeSpelling = '\002\217\000\000',
	WordWdDictionaryTypeGrammar = '\002\217\000\001',
	WordWdDictionaryTypeThesaurus = '\002\217\000\002',
	WordWdDictionaryTypeHyphenation = '\002\217\000\003',
	WordWdDictionaryTypeSpellingComplete = '\002\217\000\004',
	WordWdDictionaryTypeSpellingCustom = '\002\217\000\005',
	WordWdDictionaryTypeSpellingLegal = '\002\217\000\006',
	WordWdDictionaryTypeSpellingMedical = '\002\217\000\007',
	WordWdDictionaryTypeHangulHanjaConversion = '\002\217\000\010',
	WordWdDictionaryTypeHangulHanjaConversionCustom = '\002\217\000\011'
};
typedef enum WordWdDictionaryType WordWdDictionaryType;

enum WordWdSpellingWordType {
	WordWdSpellingWordTypeSpellingWordTypeSpellWord = '\002\220\000\000',
	WordWdSpellingWordTypeSpellingWordTypeWildcard = '\002\220\000\001',
	WordWdSpellingWordTypeSpellingWordTypeAnagram = '\002\220\000\002'
};
typedef enum WordWdSpellingWordType WordWdSpellingWordType;

enum WordWdSpellingErrorType {
	WordWdSpellingErrorTypeSpellingCorrect = '\002\221\000\000',
	WordWdSpellingErrorTypeSpellingNotInDictionary = '\002\221\000\001',
	WordWdSpellingErrorTypeSpellingCapitalization = '\002\221\000\002'
};
typedef enum WordWdSpellingErrorType WordWdSpellingErrorType;

enum WordWdProofreadingErrorType {
	WordWdProofreadingErrorTypeSpellingError = '\002\222\000\000',
	WordWdProofreadingErrorTypeGrammaticalError = '\002\222\000\001'
};
typedef enum WordWdProofreadingErrorType WordWdProofreadingErrorType;

enum WordWdInlineShapeType {
	WordWdInlineShapeTypeInlineShapeEmbeddedOleobject = '\002\223\000\001',
	WordWdInlineShapeTypeInlineShapeLinkedOleobject = '\002\223\000\002',
	WordWdInlineShapeTypeInlineShapePicture = '\002\223\000\003',
	WordWdInlineShapeTypeInlineShapeLinkedPicture = '\002\223\000\004',
	WordWdInlineShapeTypeInlineShapeOlecontrolObject = '\002\223\000\005',
	WordWdInlineShapeTypeInlineShapeHorizontalLine = '\002\223\000\006',
	WordWdInlineShapeTypeInlineShapePictureHorizontalLine = '\002\223\000\007',
	WordWdInlineShapeTypeInlineShapeLinkedPictureHorizontalLine = '\002\223\000\010',
	WordWdInlineShapeTypeInlineShapePictureBullet = '\002\223\000\011',
	WordWdInlineShapeTypeInlineShapeScriptAnchor = '\002\223\000\012',
	WordWdInlineShapeTypeInlineShapeOWSAnchor = '\002\223\000\013',
	WordWdInlineShapeTypeInlineShapeChart = '\002\223\000\014',
	WordWdInlineShapeTypeInlineShapeDiagram = '\002\223\000\015',
	WordWdInlineShapeTypeInlineShapeLockedCanvas = '\002\223\000\016',
	WordWdInlineShapeTypeInlineShapeSmartartGraphic = '\002\223\000\017',
	WordWdInlineShapeTypeInlineShapeWebVideo = '\002\223\000\020',
	WordWdInlineShapeTypeInlineShapeGraphic = '\002\223\000\021',
	WordWdInlineShapeTypeInlineShapeLinkedGraphic = '\002\223\000\022',
	WordWdInlineShapeTypeInlineShape3dmodel = '\002\223\000\023',
	WordWdInlineShapeTypeInlineShapeLinked3dmodel = '\002\223\000\024'
};
typedef enum WordWdInlineShapeType WordWdInlineShapeType;

enum WordWdArrangeStyle {
	WordWdArrangeStyleTiled = '\002\224\000\000',
	WordWdArrangeStyleIcons = '\002\224\000\001'
};
typedef enum WordWdArrangeStyle WordWdArrangeStyle;

enum WordWdSelectionFlags {
	WordWdSelectionFlagsSelectionStartActive = '\002\225\000\001',
	WordWdSelectionFlagsSelectionAtEol = '\002\225\000\002',
	WordWdSelectionFlagsSelectionOvertype = '\002\225\000\004',
	WordWdSelectionFlagsSelectionActive = '\002\225\000\010',
	WordWdSelectionFlagsSelectionReplace = '\002\225\000\020',
	WordWdSelectionFlagsSelectionInactive = '\002\225\000\000',
	WordWdSelectionFlagsSelectionStartActiveAndAtEol = '\002\225\000\003',
	WordWdSelectionFlagsSelectionStartActiveAndOvertype = '\002\225\000\005',
	WordWdSelectionFlagsSelectionAtEolAndOvertype = '\002\225\000\006',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndOvertype = '\002\225\000\007',
	WordWdSelectionFlagsSelectionStartActiveAndActive = '\002\225\000\011',
	WordWdSelectionFlagsSelectionAtEolAndActive = '\002\225\000\012',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndActive = '\002\225\000\013',
	WordWdSelectionFlagsSelectionOvertypeAndActive = '\002\225\000\014',
	WordWdSelectionFlagsSelectionStartActiveAndOvertypeAndActive = '\002\225\000\015',
	WordWdSelectionFlagsSelectionAtEolAndOvertypeAndActive = '\002\225\000\016',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndOvertypeAndActive = '\002\225\000\017',
	WordWdSelectionFlagsSelectionStartActiveAndReplace = '\002\225\000\021',
	WordWdSelectionFlagsSelectionAtEolAndReplace = '\002\225\000\022',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndReplace = '\002\225\000\023',
	WordWdSelectionFlagsSelectionOvertypeAndReplace = '\002\225\000\024',
	WordWdSelectionFlagsSelectionStartActiveAndOvertypeAndReplace = '\002\225\000\025',
	WordWdSelectionFlagsSelectionAtEolAndOvertypeAndReplace = '\002\225\000\026',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndOvertypeAndReplace = '\002\225\000\027',
	WordWdSelectionFlagsSelectionActiveAndReplace = '\002\225\000\030',
	WordWdSelectionFlagsSelectionStartActiveAndActiveAndReplace = '\002\225\000\031',
	WordWdSelectionFlagsSelectionAtEolAndActiveAndReplace = '\002\225\000\032',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndActiveAndReplace = '\002\225\000\033',
	WordWdSelectionFlagsSelectionOvertypeAndActiveAndReplace = '\002\225\000\034',
	WordWdSelectionFlagsSelectionStartActiveAndOvertypeAndActiveAndReplace = '\002\225\000\035',
	WordWdSelectionFlagsSelectionAtEolAndOvertypeAndActiveAndReplace = '\002\225\000\036',
	WordWdSelectionFlagsSelectionStartActiveAndAtEolAndOvertypeAndActiveAndReplace = '\002\225\000\037'
};
typedef enum WordWdSelectionFlags WordWdSelectionFlags;

enum WordWdAutoVersions {
	WordWdAutoVersionsAutoVersionOff = '\002\226\000\000',
	WordWdAutoVersionsAutoVersionOnClose = '\002\226\000\001'
};
typedef enum WordWdAutoVersions WordWdAutoVersions;

enum WordWdOrganizerObject {
	WordWdOrganizerObjectOrganizerObjectStyles = '\002\227\000\000',
	WordWdOrganizerObjectOrganizerObjectAutoText = '\002\227\000\001',
	WordWdOrganizerObjectOrganizerObjectCommandBars = '\002\227\000\002',
	WordWdOrganizerObjectOrganizerObjectProjectItems = '\002\227\000\003'
};
typedef enum WordWdOrganizerObject WordWdOrganizerObject;

enum WordWdFindMatch {
	WordWdFindMatchMatchParagraphMark = '\002\231\000\017',
	WordWdFindMatchMatchTabCharacter = '\002\230\000\011',
	WordWdFindMatchMatchCommentMark = '\002\230\000\005',
	WordWdFindMatchMatchAnyCharacter = '\002\231\000\?',
	WordWdFindMatchMatchAnyDigit = '\002\231\000\037',
	WordWdFindMatchMatchAnyLetter = '\002\231\000/',
	WordWdFindMatchMatchCaretCharacter = '\002\230\000\013',
	WordWdFindMatchMatchColumnBreak = '\002\230\000\016',
	WordWdFindMatchMatchEmDash = '\002\230 \024',
	WordWdFindMatchMatchEnDash = '\002\230 \023',
	WordWdFindMatchMatchEndnoteMark = '\002\231\000\023',
	WordWdFindMatchMatchField = '\002\230\000\023',
	WordWdFindMatchMatchFootnoteMark = '\002\231\000\022',
	WordWdFindMatchMatchGraphic = '\002\230\000\001',
	WordWdFindMatchMatchManualLineBreak = '\002\231\000\017',
	WordWdFindMatchMatchManualPageBreak = '\002\231\000\034',
	WordWdFindMatchMatchNonbreakingHyphen = '\002\230\000\036',
	WordWdFindMatchMatchNonbreakingSpace = '\002\230\000\240',
	WordWdFindMatchMatchOptionalHyphen = '\002\230\000\037',
	WordWdFindMatchMatchSectionBreak = '\002\231\000,',
	WordWdFindMatchMatchWhiteSpace = '\002\231\000w'
};
typedef enum WordWdFindMatch WordWdFindMatch;

enum WordWdFindWrap {
	WordWdFindWrapFindStop = '\002\231\000\000',
	WordWdFindWrapFindContinue = '\002\231\000\001',
	WordWdFindWrapFindAsk = '\002\231\000\002'
};
typedef enum WordWdFindWrap WordWdFindWrap;

enum WordWdInformation {
	WordWdInformationActiveEndAdjustedPageNumber = '\002\232\000\001',
	WordWdInformationActiveEndSectionNumber = '\002\232\000\002',
	WordWdInformationActiveEndPageNumber = '\002\232\000\003',
	WordWdInformationNumberOfPagesInDocument = '\002\232\000\004',
	WordWdInformationHorizontalPositionRelativeToPage = '\002\232\000\005',
	WordWdInformationVerticalPositionRelativeToPage = '\002\232\000\006',
	WordWdInformationHorizontalPositionRelativeToTextBoundary = '\002\232\000\007',
	WordWdInformationVerticalPositionRelativeToTextBoundary = '\002\232\000\010',
	WordWdInformationFirstCharacterColumnNumber = '\002\232\000\011',
	WordWdInformationFirstCharacterLineNumber = '\002\232\000\012',
	WordWdInformationFrameIsSelected = '\002\232\000\013',
	WordWdInformationWithInTable = '\002\232\000\014',
	WordWdInformationStartOfRangeRowNumber = '\002\232\000\015',
	WordWdInformationEnd_ofRangeRowNumber = '\002\232\000\016',
	WordWdInformationMaximumNumberOfRows = '\002\232\000\017',
	WordWdInformationStartOfRangeColumnNumber = '\002\232\000\020',
	WordWdInformationEnd_ofRangeColumnNumber = '\002\232\000\021',
	WordWdInformationMaximumNumberOfColumns = '\002\232\000\022',
	WordWdInformationZoomPercentage = '\002\232\000\023',
	WordWdInformationSelectionMode = '\002\232\000\024',
	WordWdInformationInfoCapsLock = '\002\232\000\025',
	WordWdInformationInfoNumLock = '\002\232\000\026',
	WordWdInformationOverType = '\002\232\000\027',
	WordWdInformationRevisionMarking = '\002\232\000\030',
	WordWdInformationInFootnoteEndnotePane = '\002\232\000\031',
	WordWdInformationInCommentPane = '\002\232\000\032',
	WordWdInformationInHeaderFooter = '\002\232\000\034',
	WordWdInformationAtEndOfRowMarker = '\002\232\000\037',
	WordWdInformationReferenceOfType = '\002\232\000 ',
	WordWdInformationHeaderFooterType = '\002\232\000!',
	WordWdInformationInMasterDocument = '\002\232\000\"',
	WordWdInformationInFootnote = '\002\232\000#',
	WordWdInformationInEndnote = '\002\232\000$',
	WordWdInformationInWordMail = '\002\232\000%',
	WordWdInformationInClipboard = '\002\232\000&',
	WordWdInformationInCoverPage = '\002\232\000)',
	WordWdInformationInBibliography = '\002\232\000*',
	WordWdInformationInCitation = '\002\232\000+',
	WordWdInformationInFieldCode = '\002\232\000,',
	WordWdInformationInFieldResult = '\002\232\000-',
	WordWdInformationInContentControl = '\002\232\000.'
};
typedef enum WordWdInformation WordWdInformation;

enum WordWdWrapType {
	WordWdWrapTypeWrapSquare = '\002\233\000\000',
	WordWdWrapTypeWrapTight = '\002\233\000\001',
	WordWdWrapTypeWrapThrough = '\002\233\000\002',
	WordWdWrapTypeWrapNone = '\002\233\000\003',
	WordWdWrapTypeWrapTopBottom = '\002\233\000\004',
	WordWdWrapTypeWrapBehind = '\002\233\000\005',
	WordWdWrapTypeWrapFront = '\002\233\000\003',
	WordWdWrapTypeWrapInline = '\002\233\000\007'
};
typedef enum WordWdWrapType WordWdWrapType;

enum WordWdWrapTypeMerged {
	WordWdWrapTypeMergedPictureWrapTypeInlineWithText = '\002\310\000\000',
	WordWdWrapTypeMergedPictureWrapTypeSquare = '\002\310\000\001',
	WordWdWrapTypeMergedPictureWrapTypeTight = '\002\310\000\002',
	WordWdWrapTypeMergedPictureWrapTypeBehindText = '\002\310\000\003',
	WordWdWrapTypeMergedPictureWrapTypeInFrontOfText = '\002\310\000\004',
	WordWdWrapTypeMergedPictureWrapTypeThrough = '\002\310\000\005',
	WordWdWrapTypeMergedPictureWrapTypeTopAndBottom = '\002\310\000\006'
};
typedef enum WordWdWrapTypeMerged WordWdWrapTypeMerged;

enum WordWdWrapSideType {
	WordWdWrapSideTypeWrapBoth = '\002\234\000\000',
	WordWdWrapSideTypeWrapLeft = '\002\234\000\001',
	WordWdWrapSideTypeWrapRight = '\002\234\000\002',
	WordWdWrapSideTypeWrapLargest = '\002\234\000\003'
};
typedef enum WordWdWrapSideType WordWdWrapSideType;

enum WordWdOutlineLevel {
	WordWdOutlineLevelOutlineLevel1 = '\002\235\000\001',
	WordWdOutlineLevelOutlineLevel2 = '\002\235\000\002',
	WordWdOutlineLevelOutlineLevel3 = '\002\235\000\003',
	WordWdOutlineLevelOutlineLevel4 = '\002\235\000\004',
	WordWdOutlineLevelOutlineLevel5 = '\002\235\000\005',
	WordWdOutlineLevelOutlineLevel6 = '\002\235\000\006',
	WordWdOutlineLevelOutlineLevel7 = '\002\235\000\007',
	WordWdOutlineLevelOutlineLevel8 = '\002\235\000\010',
	WordWdOutlineLevelOutlineLevel9 = '\002\235\000\011',
	WordWdOutlineLevelOutlineLevelBodyText = '\002\235\000\012'
};
typedef enum WordWdOutlineLevel WordWdOutlineLevel;

enum WordWdTextOrientation {
	WordWdTextOrientationTextOrientationHorizontal = '\002\236\000\000',
	WordWdTextOrientationTextOrientationUpward = '\002\236\000\002',
	WordWdTextOrientationTextOrientationDownward = '\002\236\000\003',
	WordWdTextOrientationTextOrientationVerticalEastAsian = '\002\236\000\001',
	WordWdTextOrientationTextOrientationHorizontalRotatedEastAsian = '\002\236\000\004',
	WordWdTextOrientationTextOrientationVertical = '\002\236\000\005'
};
typedef enum WordWdTextOrientation WordWdTextOrientation;

enum WordWdPageBorderArt {
	WordWdPageBorderArtArtNone = '\002\237\000\000',
	WordWdPageBorderArtArtApples = '\002\237\000\001',
	WordWdPageBorderArtArtMapleMuffins = '\002\237\000\002',
	WordWdPageBorderArtArtCakeSlice = '\002\237\000\003',
	WordWdPageBorderArtArtCandyCorn = '\002\237\000\004',
	WordWdPageBorderArtArtIceCreamCones = '\002\237\000\005',
	WordWdPageBorderArtArtChampagneBottle = '\002\237\000\006',
	WordWdPageBorderArtArtPartyGlass = '\002\237\000\007',
	WordWdPageBorderArtArtChristmasTree = '\002\237\000\010',
	WordWdPageBorderArtArtTrees = '\002\237\000\011',
	WordWdPageBorderArtArtPalmsColor = '\002\237\000\012',
	WordWdPageBorderArtArtBalloons3Colors = '\002\237\000\013',
	WordWdPageBorderArtArtBalloonsHotAir = '\002\237\000\014',
	WordWdPageBorderArtArtPartyFavor = '\002\237\000\015',
	WordWdPageBorderArtArtConfettiStreamers = '\002\237\000\016',
	WordWdPageBorderArtArtHearts = '\002\237\000\017',
	WordWdPageBorderArtArtHeartBalloon = '\002\237\000\020',
	WordWdPageBorderArtArtStars3D = '\002\237\000\021',
	WordWdPageBorderArtArtStarsShadowed = '\002\237\000\022',
	WordWdPageBorderArtArtStars = '\002\237\000\023',
	WordWdPageBorderArtArtSun = '\002\237\000\024',
	WordWdPageBorderArtArtEarth2 = '\002\237\000\025',
	WordWdPageBorderArtArtEarth1 = '\002\237\000\026',
	WordWdPageBorderArtArtPeopleHats = '\002\237\000\027',
	WordWdPageBorderArtArtSombrero = '\002\237\000\030',
	WordWdPageBorderArtArtPencils = '\002\237\000\031',
	WordWdPageBorderArtArtPackages = '\002\237\000\032',
	WordWdPageBorderArtArtClocks = '\002\237\000\033',
	WordWdPageBorderArtArtFirecrackers = '\002\237\000\034',
	WordWdPageBorderArtArtRings = '\002\237\000\035',
	WordWdPageBorderArtArtMapPins = '\002\237\000\036',
	WordWdPageBorderArtArtConfetti = '\002\237\000\037',
	WordWdPageBorderArtArtCreaturesButterfly = '\002\237\000 ',
	WordWdPageBorderArtArtCreaturesLadyBug = '\002\237\000!',
	WordWdPageBorderArtArtCreaturesFish = '\002\237\000\"',
	WordWdPageBorderArtArtBirdsFlight = '\002\237\000#',
	WordWdPageBorderArtArtScaredCat = '\002\237\000$',
	WordWdPageBorderArtArtBats = '\002\237\000%',
	WordWdPageBorderArtArtFlowersRoses = '\002\237\000&',
	WordWdPageBorderArtArtFlowersRedRose = '\002\237\000\'',
	WordWdPageBorderArtArtPoinsettias = '\002\237\000(',
	WordWdPageBorderArtArtHolly = '\002\237\000)',
	WordWdPageBorderArtArtFlowersTiny = '\002\237\000*',
	WordWdPageBorderArtArtFlowersPansy = '\002\237\000+',
	WordWdPageBorderArtArtFlowersModern2 = '\002\237\000,',
	WordWdPageBorderArtArtFlowersModern1 = '\002\237\000-',
	WordWdPageBorderArtArtWhiteFlowers = '\002\237\000.',
	WordWdPageBorderArtArtVine = '\002\237\000/',
	WordWdPageBorderArtArtFlowersDaisies = '\002\237\0000',
	WordWdPageBorderArtArtFlowersBlockPrint = '\002\237\0001',
	WordWdPageBorderArtArtDecoArchColor = '\002\237\0002',
	WordWdPageBorderArtArtFans = '\002\237\0003',
	WordWdPageBorderArtArtFilm = '\002\237\0004',
	WordWdPageBorderArtArtLightning1 = '\002\237\0005',
	WordWdPageBorderArtArtCompass = '\002\237\0006',
	WordWdPageBorderArtArtDoubleD = '\002\237\0007',
	WordWdPageBorderArtArtClassicalWave = '\002\237\0008',
	WordWdPageBorderArtArtShadowedSquares = '\002\237\0009',
	WordWdPageBorderArtArtTwistedLines1 = '\002\237\000:',
	WordWdPageBorderArtArtWaveline = '\002\237\000;',
	WordWdPageBorderArtArtQuadrants = '\002\237\000<',
	WordWdPageBorderArtArtCheckedBarColor = '\002\237\000=',
	WordWdPageBorderArtArtSwirligig = '\002\237\000>',
	WordWdPageBorderArtArtPushPinNote1 = '\002\237\000\?',
	WordWdPageBorderArtArtPushPinNote2 = '\002\237\000@',
	WordWdPageBorderArtArtPumpkin1 = '\002\237\000A',
	WordWdPageBorderArtArtEggsBlack = '\002\237\000B',
	WordWdPageBorderArtArtCup = '\002\237\000C',
	WordWdPageBorderArtArtHeartGray = '\002\237\000D',
	WordWdPageBorderArtArtGingerbreadMan = '\002\237\000E',
	WordWdPageBorderArtArtBabyPacifier = '\002\237\000F',
	WordWdPageBorderArtArtBabyRattle = '\002\237\000G',
	WordWdPageBorderArtArtCabins = '\002\237\000H',
	WordWdPageBorderArtArtHouseFunky = '\002\237\000I',
	WordWdPageBorderArtArtStarsBlack = '\002\237\000J',
	WordWdPageBorderArtArtSnowflakes = '\002\237\000K',
	WordWdPageBorderArtArtSnowflakeFancy = '\002\237\000L',
	WordWdPageBorderArtArtSkyrocket = '\002\237\000M',
	WordWdPageBorderArtArtSeattle = '\002\237\000N',
	WordWdPageBorderArtArtMusicNotes = '\002\237\000O',
	WordWdPageBorderArtArtPalmsBlack = '\002\237\000P',
	WordWdPageBorderArtArtMapleLeaf = '\002\237\000Q',
	WordWdPageBorderArtArtPaperClips = '\002\237\000R',
	WordWdPageBorderArtArtShorebirdTracks = '\002\237\000S',
	WordWdPageBorderArtArtPeople = '\002\237\000T',
	WordWdPageBorderArtArtPeopleWaving = '\002\237\000U',
	WordWdPageBorderArtArtEclipsingSquares2 = '\002\237\000V',
	WordWdPageBorderArtArtHypnotic = '\002\237\000W',
	WordWdPageBorderArtArtDiamondsGray = '\002\237\000X',
	WordWdPageBorderArtArtDecoArch = '\002\237\000Y',
	WordWdPageBorderArtArtDecoBlocks = '\002\237\000Z',
	WordWdPageBorderArtArtCirclesLines = '\002\237\000[',
	WordWdPageBorderArtArtPapyrus = '\002\237\000\\',
	WordWdPageBorderArtArtWoodwork = '\002\237\000]',
	WordWdPageBorderArtArtWeavingBraid = '\002\237\000^',
	WordWdPageBorderArtArtWeavingRibbon = '\002\237\000_',
	WordWdPageBorderArtArtWeavingAngles = '\002\237\000`',
	WordWdPageBorderArtArtArchedScallops = '\002\237\000a',
	WordWdPageBorderArtArtSafari = '\002\237\000b',
	WordWdPageBorderArtArtCelticKnotwork = '\002\237\000c',
	WordWdPageBorderArtArtCrazyMaze = '\002\237\000d',
	WordWdPageBorderArtArtEclipsingSquares1 = '\002\237\000e',
	WordWdPageBorderArtArtBirds = '\002\237\000f',
	WordWdPageBorderArtArtFlowersTeacup = '\002\237\000g',
	WordWdPageBorderArtArtNorthwest = '\002\237\000h',
	WordWdPageBorderArtArtSouthwest = '\002\237\000i',
	WordWdPageBorderArtArtTribal6 = '\002\237\000j',
	WordWdPageBorderArtArtTribal4 = '\002\237\000k',
	WordWdPageBorderArtArtTribal3 = '\002\237\000l',
	WordWdPageBorderArtArtTribal2 = '\002\237\000m',
	WordWdPageBorderArtArtTribal5 = '\002\237\000n',
	WordWdPageBorderArtArtXillusions = '\002\237\000o',
	WordWdPageBorderArtArtZanyTriangles = '\002\237\000p',
	WordWdPageBorderArtArtPyramids = '\002\237\000q',
	WordWdPageBorderArtArtPyramidsAbove = '\002\237\000r',
	WordWdPageBorderArtArtConfettiGrays = '\002\237\000s',
	WordWdPageBorderArtArtConfettiOutline = '\002\237\000t',
	WordWdPageBorderArtArtConfettiWhite = '\002\237\000u',
	WordWdPageBorderArtArtMosaic = '\002\237\000v',
	WordWdPageBorderArtArtLightning2 = '\002\237\000w',
	WordWdPageBorderArtArtHeebieJeebies = '\002\237\000x',
	WordWdPageBorderArtArtLightBulb = '\002\237\000y',
	WordWdPageBorderArtArtGradient = '\002\237\000z',
	WordWdPageBorderArtArtTriangleParty = '\002\237\000{',
	WordWdPageBorderArtArtTwistedLines2 = '\002\237\000|',
	WordWdPageBorderArtArtMoons = '\002\237\000}',
	WordWdPageBorderArtArtOvals = '\002\237\000~',
	WordWdPageBorderArtArtDoubleDiamonds = '\002\237\000\177',
	WordWdPageBorderArtArtChainLink = '\002\237\000\200',
	WordWdPageBorderArtArtTriangles = '\002\237\000\201',
	WordWdPageBorderArtArtTribal1 = '\002\237\000\202',
	WordWdPageBorderArtArtMarqueeToothed = '\002\237\000\203',
	WordWdPageBorderArtArtSharksTeeth = '\002\237\000\204',
	WordWdPageBorderArtArtSawtooth = '\002\237\000\205',
	WordWdPageBorderArtArtSawtoothGray = '\002\237\000\206',
	WordWdPageBorderArtArtPostageStamp = '\002\237\000\207',
	WordWdPageBorderArtArtWeavingStrips = '\002\237\000\210',
	WordWdPageBorderArtArtZigZag = '\002\237\000\211',
	WordWdPageBorderArtArtCrossStitch = '\002\237\000\212',
	WordWdPageBorderArtArtGems = '\002\237\000\213',
	WordWdPageBorderArtArtCirclesRectangles = '\002\237\000\214',
	WordWdPageBorderArtArtCornerTriangles = '\002\237\000\215',
	WordWdPageBorderArtArtCreaturesInsects = '\002\237\000\216',
	WordWdPageBorderArtArtZigZagStitch = '\002\237\000\217',
	WordWdPageBorderArtArtCheckered = '\002\237\000\220',
	WordWdPageBorderArtArtCheckedBarBlack = '\002\237\000\221',
	WordWdPageBorderArtArtMarquee = '\002\237\000\222',
	WordWdPageBorderArtArtBasicWhiteDots = '\002\237\000\223',
	WordWdPageBorderArtArtBasicWideMidline = '\002\237\000\224',
	WordWdPageBorderArtArtBasicWideOutline = '\002\237\000\225',
	WordWdPageBorderArtArtBasicWideInline = '\002\237\000\226',
	WordWdPageBorderArtArtBasicThinLines = '\002\237\000\227',
	WordWdPageBorderArtArtBasicWhiteDashes = '\002\237\000\230',
	WordWdPageBorderArtArtBasicWhiteSquares = '\002\237\000\231',
	WordWdPageBorderArtArtBasicBlackSquares = '\002\237\000\232',
	WordWdPageBorderArtArtBasicBlackDashes = '\002\237\000\233',
	WordWdPageBorderArtArtBasicBlackDots = '\002\237\000\234',
	WordWdPageBorderArtArtStarsTop = '\002\237\000\235',
	WordWdPageBorderArtArtCertificateBanner = '\002\237\000\236',
	WordWdPageBorderArtArtHandmade1 = '\002\237\000\237',
	WordWdPageBorderArtArtHandmade2 = '\002\237\000\240',
	WordWdPageBorderArtArtTornPaper = '\002\237\000\241',
	WordWdPageBorderArtArtTornPaperBlack = '\002\237\000\242',
	WordWdPageBorderArtArtCouponCutoutDashes = '\002\237\000\243',
	WordWdPageBorderArtArtCouponCutoutDots = '\002\237\000\244'
};
typedef enum WordWdPageBorderArt WordWdPageBorderArt;

enum WordWdBorderDistanceFrom {
	WordWdBorderDistanceFromBorderDistanceFromText = '\002\240\000\000',
	WordWdBorderDistanceFromBorderDistanceFromPageEdge = '\002\240\000\001'
};
typedef enum WordWdBorderDistanceFrom WordWdBorderDistanceFrom;

enum WordWdReplace {
	WordWdReplaceReplaceNone = '\002\241\000\000',
	WordWdReplaceReplaceOne = '\002\241\000\001',
	WordWdReplaceReplaceAll = '\002\241\000\002'
};
typedef enum WordWdReplace WordWdReplace;

enum WordWdFontBias {
	WordWdFontBiasFontBiasDoNotCare = '\002\242\000\377',
	WordWdFontBiasFontBiasDefault = '\002\242\000\000',
	WordWdFontBiasFontBiasEastAsian = '\002\242\000\001'
};
typedef enum WordWdFontBias WordWdFontBias;

enum WordWdBrowserLevel {
	WordWdBrowserLevelBrowserLevelV4 = '\002\243\000\000',
	WordWdBrowserLevelBrowserLevelMicrosoftInternetExplorer5 = '\002\243\000\001',
	WordWdBrowserLevelBrowserLevelMicrosoftInternetExplorer6 = '\002\243\000\002'
};
typedef enum WordWdBrowserLevel WordWdBrowserLevel;

enum WordWdEnclosureType {
	WordWdEnclosureTypeEnclosureCircle = '\002\244\000\000',
	WordWdEnclosureTypeEnclosureSquare = '\002\244\000\001',
	WordWdEnclosureTypeEnclosureTriangle = '\002\244\000\002',
	WordWdEnclosureTypeEnclosureDiamond = '\002\244\000\003'
};
typedef enum WordWdEnclosureType WordWdEnclosureType;

enum WordWdEncloseStyle {
	WordWdEncloseStyleEncloseStyleNone = '\002\245\000\000',
	WordWdEncloseStyleEncloseStyleSmall = '\002\245\000\001',
	WordWdEncloseStyleEncloseStyleLarge = '\002\245\000\002'
};
typedef enum WordWdEncloseStyle WordWdEncloseStyle;

enum WordWdLayoutMode {
	WordWdLayoutModeLayoutModeDefault = '\002\246\000\000',
	WordWdLayoutModeLayoutModeGrid = '\002\246\000\001',
	WordWdLayoutModeLayoutModeLineGrid = '\002\246\000\002',
	WordWdLayoutModeLayoutModeGenko = '\002\246\000\003'
};
typedef enum WordWdLayoutMode WordWdLayoutMode;

enum WordWdDocumentMedium {
	WordWdDocumentMediumForAEmailMessage = '\002\247\000\000',
	WordWdDocumentMediumForADocument = '\002\247\000\001',
	WordWdDocumentMediumForAWebPage = '\002\247\000\002'
};
typedef enum WordWdDocumentMedium WordWdDocumentMedium;

enum WordWdTwoLinesInOneType {
	WordWdTwoLinesInOneTypeTwoLinesInOneNone = '\002\304\000\000',
	WordWdTwoLinesInOneTypeTwoLinesInOneNoBrackets = '\002\304\000\001',
	WordWdTwoLinesInOneTypeTwoLinesInOneParentheses = '\002\304\000\002',
	WordWdTwoLinesInOneTypeTwoLinesInOneSquareBrackets = '\002\304\000\003',
	WordWdTwoLinesInOneTypeTwoLinesInOneAngleBrackets = '\002\304\000\004',
	WordWdTwoLinesInOneTypeTwoLinesInOneCurlyBrackets = '\002\304\000\005'
};
typedef enum WordWdTwoLinesInOneType WordWdTwoLinesInOneType;

enum WordWdHorizontalInVerticalType {
	WordWdHorizontalInVerticalTypeHorizontalInVerticalNone = '\002\305\000\000',
	WordWdHorizontalInVerticalTypeHorizontalInVerticalFitInLine = '\002\305\000\001',
	WordWdHorizontalInVerticalTypeHorizontalInVerticalResizeLine = '\002\305\000\002'
};
typedef enum WordWdHorizontalInVerticalType WordWdHorizontalInVerticalType;

enum WordWdHorizontalLineAlignment {
	WordWdHorizontalLineAlignmentHorizontalLineAlignLeft = '\002\250\000\000',
	WordWdHorizontalLineAlignmentHorizontalLineAlignCenter = '\002\250\000\001',
	WordWdHorizontalLineAlignmentHorizontalLineAlignRight = '\002\250\000\002'
};
typedef enum WordWdHorizontalLineAlignment WordWdHorizontalLineAlignment;

enum WordWdHorizontalLineWidthType {
	WordWdHorizontalLineWidthTypeHorizontalLinePercentWidth = '\002\250\377\377',
	WordWdHorizontalLineWidthTypeHorizontalLineFixedWidth = '\002\250\377\376'
};
typedef enum WordWdHorizontalLineWidthType WordWdHorizontalLineWidthType;

enum WordWdPhoneticGuideAlignmentType {
	WordWdPhoneticGuideAlignmentTypePhoneticGuideAlignmentCenter = '\002\306\000\000',
	WordWdPhoneticGuideAlignmentTypePhoneticGuideAlignmentZeroOneZero = '\002\306\000\001',
	WordWdPhoneticGuideAlignmentTypePhoneticGuideAlignmentOneTwoOne = '\002\306\000\002',
	WordWdPhoneticGuideAlignmentTypePhoneticGuideAlignmentLeft = '\002\306\000\003',
	WordWdPhoneticGuideAlignmentTypePhoneticGuideAlignmentRight = '\002\306\000\004',
	WordWdPhoneticGuideAlignmentTypePhoneticGuideAlignmentRightVertical = '\002\306\000\005'
};
typedef enum WordWdPhoneticGuideAlignmentType WordWdPhoneticGuideAlignmentType;

enum WordWdNumberStyleWordBasicBiDi {
	WordWdNumberStyleWordBasicBiDiListNumberStyleBidi1 = '\002\335\0001',
	WordWdNumberStyleWordBasicBiDiListNumberStyleBidi2 = '\002\335\0002',
	WordWdNumberStyleWordBasicBiDiCaptionNumberStyleBidiLetter1 = '\002\335\0001',
	WordWdNumberStyleWordBasicBiDiCaptionNumberStyleBidiLetter2 = '\002\335\0002',
	WordWdNumberStyleWordBasicBiDiNoteNumberStyleBidiLetter1 = '\002\335\0001',
	WordWdNumberStyleWordBasicBiDiNoteNumberStyleBidiLetter2 = '\002\335\0002',
	WordWdNumberStyleWordBasicBiDiPageNumberStyleBidiLetter1 = '\002\335\0001',
	WordWdNumberStyleWordBasicBiDiPageNumberStyleBidiLetter2 = '\002\335\0002'
};
typedef enum WordWdNumberStyleWordBasicBiDi WordWdNumberStyleWordBasicBiDi;

enum WordWdTableDirection {
	WordWdTableDirectionTableDirectionRtl = '\002\252\000\000',
	WordWdTableDirectionTableDirectionLtr = '\002\252\000\001'
};
typedef enum WordWdTableDirection WordWdTableDirection;

enum WordWdGutterStyle {
	WordWdGutterStyleGutterPositionLeft = '\002\253\000\000',
	WordWdGutterStyleGutterPositionTop = '\002\253\000\001',
	WordWdGutterStyleGutterPositionRight = '\002\253\000\002'
};
typedef enum WordWdGutterStyle WordWdGutterStyle;

enum WordWdGutterStyleOld {
	WordWdGutterStyleOldGutterStyleLatin = '\002\253\377\366',
	WordWdGutterStyleOldGutterStyleBidi = '\002\254\000\002',
	WordWdGutterStyleOldGutterStyleNone = '\002\254\000\000'
};
typedef enum WordWdGutterStyleOld WordWdGutterStyleOld;

enum WordWdShapePosition {
	WordWdShapePositionShapeTop = '\002\235\275\301',
	WordWdShapePositionShapeLeft = '\002\235\275\302',
	WordWdShapePositionShapeBottom = '\002\235\275\303',
	WordWdShapePositionShapeRight = '\002\235\275\304',
	WordWdShapePositionShapeCenter = '\002\235\275\305',
	WordWdShapePositionShapeInside = '\002\235\275\306',
	WordWdShapePositionShapeOutside = '\002\235\275\307'
};
typedef enum WordWdShapePosition WordWdShapePosition;

enum WordWdTablePosition {
	WordWdTablePositionTableTop = '\002\236\275\301',
	WordWdTablePositionTableLeft = '\002\236\275\302',
	WordWdTablePositionTableBottom = '\002\236\275\303',
	WordWdTablePositionTableRight = '\002\236\275\304',
	WordWdTablePositionTableCenter = '\002\236\275\305',
	WordWdTablePositionTableInside = '\002\236\275\306',
	WordWdTablePositionTableOutside = '\002\236\275\307'
};
typedef enum WordWdTablePosition WordWdTablePosition;

enum WordWdDefaultTableBehavior {
	WordWdDefaultTableBehaviorWord8TableBehavior = '\002\257\000\000',
	WordWdDefaultTableBehaviorWord9TableBehavior = '\002\257\000\001'
};
typedef enum WordWdDefaultTableBehavior WordWdDefaultTableBehavior;

enum WordWdAutoFitBehavior {
	WordWdAutoFitBehaviorAutoFitFixed = '\002\260\000\000',
	WordWdAutoFitBehaviorAutoFitContent = '\002\260\000\001',
	WordWdAutoFitBehaviorAutoFitWindow = '\002\260\000\002'
};
typedef enum WordWdAutoFitBehavior WordWdAutoFitBehavior;

enum WordWdDefaultListBehavior {
	WordWdDefaultListBehaviorWord8ListBehavior = '\002\261\000\000',
	WordWdDefaultListBehaviorWord9ListBehavior = '\002\261\000\001',
	WordWdDefaultListBehaviorWord10ListBehavior = '\002\261\000\002'
};
typedef enum WordWdDefaultListBehavior WordWdDefaultListBehavior;

enum WordWdPreferredWidthType {
	WordWdPreferredWidthTypePreferredWidthAuto = '\002\262\000\001',
	WordWdPreferredWidthTypePreferredWidthPercent = '\002\262\000\002',
	WordWdPreferredWidthTypePreferredWidthPoints = '\002\262\000\003'
};
typedef enum WordWdPreferredWidthType WordWdPreferredWidthType;

enum WordWdNewDocumentType {
	WordWdNewDocumentTypeNewBlankDocument = '\002\263\000\000',
	WordWdNewDocumentTypeNewWebPage = '\002\263\000\001',
	WordWdNewDocumentTypeNewXmlDocument = '\002\263\000\004'
};
typedef enum WordWdNewDocumentType WordWdNewDocumentType;

enum WordWdRecoveryType {
	WordWdRecoveryTypePasteDefault = '\002\265\000\000',
	WordWdRecoveryTypeSingleCellText = '\002\265\000\005',
	WordWdRecoveryTypeSingleCellTable = '\002\265\000\006',
	WordWdRecoveryTypeListContinueNumbering = '\002\265\000\007',
	WordWdRecoveryTypeListRestartNumbering = '\002\265\000\010',
	WordWdRecoveryTypeTableInsertAsRows = '\002\265\000\013',
	WordWdRecoveryTypeTableAppendTable = '\002\265\000\012',
	WordWdRecoveryTypeTableOriginalFormatting = '\002\265\000\014',
	WordWdRecoveryTypeChartPicture = '\002\265\000\015',
	WordWdRecoveryTypeChart = '\002\265\000\016',
	WordWdRecoveryTypeChartLinked = '\002\265\000\017',
	WordWdRecoveryTypeFormatOriginalFormatting = '\002\265\000\020',
	WordWdRecoveryTypeFormatSurroundingFormattingWithEmphasis = '\002\265\000\024',
	WordWdRecoveryTypeFormatPlainText = '\002\265\000\026',
	WordWdRecoveryTypeTableOverwriteCells = '\002\265\000\027',
	WordWdRecoveryTypeListCombineWithExistingList = '\002\265\000\030',
	WordWdRecoveryTypeListDontMerge = '\002\265\000\031',
	WordWdRecoveryTypeUseDestinationStylesRecovery = '\002\265\000\023'
};
typedef enum WordWdRecoveryType WordWdRecoveryType;

enum WordWdLineEndingType {
	WordWdLineEndingTypeLineEndingCrLf = '\002\307\000\000',
	WordWdLineEndingTypeLineEndingCrOnly = '\002\307\000\001',
	WordWdLineEndingTypeLineEndingLfOnly = '\002\307\000\002',
	WordWdLineEndingTypeLineEndingLfCr = '\002\307\000\003',
	WordWdLineEndingTypeLineEndingLsPs = '\002\307\000\004'
};
typedef enum WordWdLineEndingType WordWdLineEndingType;

enum WordWdConditionCode {
	WordWdConditionCodeConditionFirstRow = '\002\301\000\000',
	WordWdConditionCodeConditionLastRow = '\002\301\000\001',
	WordWdConditionCodeConditionOddRowBanding = '\002\301\000\002',
	WordWdConditionCodeConditionEvenRowBanding = '\002\301\000\003',
	WordWdConditionCodeConditionFirstColumn = '\002\301\000\004',
	WordWdConditionCodeConditionLastColumn = '\002\301\000\005',
	WordWdConditionCodeConditionOddColumnBanding = '\002\301\000\006',
	WordWdConditionCodeConditionEvenColumnBanding = '\002\301\000\007',
	WordWdConditionCodeConditionTopRightCell = '\002\301\000\010',
	WordWdConditionCodeConditionTopLeftCell = '\002\301\000\011',
	WordWdConditionCodeConditionBottomRightCell = '\002\301\000\012',
	WordWdConditionCodeConditionBottomLeftCell = '\002\301\000\013'
};
typedef enum WordWdConditionCode WordWdConditionCode;

enum WordWdHighlightStuff {
	WordWdHighlightStuffHighlightOn = '\002\270\377\377',
	WordWdHighlightStuffHighlightOff = '\002\270\000\000',
	WordWdHighlightStuffANumericConstant = '\002\270\000\000'
};
typedef enum WordWdHighlightStuff WordWdHighlightStuff;

enum WordWdCompareTarget {
	WordWdCompareTargetCompareTargetSelected = '\002\271\000\000',
	WordWdCompareTargetCompareTargetCurrent = '\002\271\000\001',
	WordWdCompareTargetCompareTargetNew = '\002\271\000\002'
};
typedef enum WordWdCompareTarget WordWdCompareTarget;

enum WordWdMergeTarget {
	WordWdMergeTargetMergeTargetSelected = '\002\272\000\000',
	WordWdMergeTargetMergeTargetCurrent = '\002\272\000\001',
	WordWdMergeTargetMergeTargetNew = '\002\272\000\002'
};
typedef enum WordWdMergeTarget WordWdMergeTarget;

enum WordWdRevisionsView {
	WordWdRevisionsViewRevisionsViewFinal = '\002\274\000\000',
	WordWdRevisionsViewRevisionsViewOriginal = '\002\274\000\001'
};
typedef enum WordWdRevisionsView WordWdRevisionsView;

enum WordWdRevisionsMode {
	WordWdRevisionsModeBalloonRevisions = '\002\275\000\000',
	WordWdRevisionsModeInLineRevisions = '\002\275\000\001',
	WordWdRevisionsModeMixedRevisions = '\002\275\000\002'
};
typedef enum WordWdRevisionsMode WordWdRevisionsMode;

enum WordWdRevisionsBalloonPrintOrientation {
	WordWdRevisionsBalloonPrintOrientationBalloonPrintOrientationAutomatic = '\002\277\000\000',
	WordWdRevisionsBalloonPrintOrientationBalloonPrintOrientationPreserve = '\002\277\000\001',
	WordWdRevisionsBalloonPrintOrientationBalloonPrintOrientationLandscape = '\002\277\000\002'
};
typedef enum WordWdRevisionsBalloonPrintOrientation WordWdRevisionsBalloonPrintOrientation;

enum WordWdRevisionsBalloonMargin {
	WordWdRevisionsBalloonMarginBalloonMarginLeft = '\002\300\000\000',
	WordWdRevisionsBalloonMarginBalloonMarginRight = '\002\300\000\001'
};
typedef enum WordWdRevisionsBalloonMargin WordWdRevisionsBalloonMargin;

enum WordWdCheckInVersionType {
	WordWdCheckInVersionTypeMinorVersion = '\002\333\000\000',
	WordWdCheckInVersionTypeMajorVersion = '\002\333\000\001',
	WordWdCheckInVersionTypeOverwriteCurrentVersion = '\002\333\000\002'
};
typedef enum WordWdCheckInVersionType WordWdCheckInVersionType;

enum WordWdLockType {
	WordWdLockTypeLockNone = '\002\334\000\000',
	WordWdLockTypeLockReservation = '\002\334\000\001',
	WordWdLockTypeLockEphemeral = '\002\334\000\002',
	WordWdLockTypeLockChanged = '\002\334\000\003'
};
typedef enum WordWdLockType WordWdLockType;

enum WordUnlink {
	WordUnlinkFieldOptions = 'w298',
	WordUnlinkField = 'w170'
};
typedef enum WordUnlink WordUnlink;

enum WordUpdateSource {
	WordUpdateSourceFieldOptions = 'w298',
	WordUpdateSourceField = 'w170'
};
typedef enum WordUpdateSource WordUpdateSource;

enum WordAccept {
	WordAcceptRevision = 'w219',
	WordAcceptConflict = 'o120'
};
typedef enum WordAccept WordAccept;

enum WordReject {
	WordRejectRevision = 'w219',
	WordRejectConflict = 'o120'
};
typedef enum WordReject WordReject;

enum WordResetSeparator {
	WordResetSeparatorFootnoteOptions = 'WopN',
	WordResetSeparatorEndnoteOptions = 'WopE'
};
typedef enum WordResetSeparator WordResetSeparator;

enum WordResetContinuationSeparator {
	WordResetContinuationSeparatorFootnoteOptions = 'WopN',
	WordResetContinuationSeparatorEndnoteOptions = 'WopE'
};
typedef enum WordResetContinuationSeparator WordResetContinuationSeparator;

enum WordResetContinuationNotice {
	WordResetContinuationNoticeFootnoteOptions = 'WopN',
	WordResetContinuationNoticeEndnoteOptions = 'WopE'
};
typedef enum WordResetContinuationNotice WordResetContinuationNotice;

enum WordActivateObject {
	WordActivateObjectDocument = 'docu',
	WordActivateObjectWindow = 'cwin',
	WordActivateObjectPane = 'w120'
};
typedef enum WordActivateObject WordActivateObject;

enum WordGetBorder {
	WordGetBorderFont = 'w137',
	WordGetBorderFrame = 'w175',
	WordGetBorderSelectionObject = 'WSoj'
};
typedef enum WordGetBorder WordGetBorder;

enum WordCutObject {
	WordCutObjectField = 'w170',
	WordCutObjectFrame = 'w175',
	WordCutObjectFormField = 'w177',
	WordCutObjectDataMergeField = 'w187',
	WordCutObjectSelectionObject = 'WSoj',
	WordCutObjectPageNumber = 'w225'
};
typedef enum WordCutObject WordCutObject;

enum WordCanContinuePreviousList {
	WordCanContinuePreviousListListFormat = 'w123',
	WordCanContinuePreviousListWordList = 'w236'
};
typedef enum WordCanContinuePreviousList WordCanContinuePreviousList;

enum WordCheckGrammar {
	WordCheckGrammarApplication = 'capp',
	WordCheckGrammarDocument = 'docu'
};
typedef enum WordCheckGrammar WordCheckGrammar;

enum WordCheckSpelling {
	WordCheckSpellingApplication = 'capp',
	WordCheckSpellingDocument = 'docu'
};
typedef enum WordCheckSpelling WordCheckSpelling;

enum WordClearFormatting {
	WordClearFormattingFind = 'w124',
	WordClearFormattingReplacement = 'w125',
	WordClearFormattingSelectionObject = 'WSoj'
};
typedef enum WordClearFormatting WordClearFormatting;

enum WordClear {
	WordClearDropCap = 'w133',
	WordClearTabStop = 'w135',
	WordClearTextInput = 'w178',
	WordClearKeyBinding = 'w242'
};
typedef enum WordClear WordClear;

enum WordCountNumberedItems {
	WordCountNumberedItemsDocument = 'docu',
	WordCountNumberedItemsListFormat = 'w123',
	WordCountNumberedItemsWordList = 'w236'
};
typedef enum WordCountNumberedItems WordCountNumberedItems;

enum WordCopyObject {
	WordCopyObjectField = 'w170',
	WordCopyObjectFrame = 'w175',
	WordCopyObjectFormField = 'w177',
	WordCopyObjectDataMergeField = 'w187',
	WordCopyObjectSelectionObject = 'WSoj',
	WordCopyObjectPageNumber = 'w225'
};
typedef enum WordCopyObject WordCopyObject;

enum WordLargeScroll {
	WordLargeScrollWindow = 'cwin',
	WordLargeScrollPane = 'w120'
};
typedef enum WordLargeScroll WordLargeScroll;

enum WordConvertNumbersToText {
	WordConvertNumbersToTextDocument = 'docu',
	WordConvertNumbersToTextListFormat = 'w123',
	WordConvertNumbersToTextWordList = 'w236'
};
typedef enum WordConvertNumbersToText WordConvertNumbersToText;

enum WordPageScroll {
	WordPageScrollWindow = 'cwin',
	WordPageScrollPane = 'w120'
};
typedef enum WordPageScroll WordPageScroll;

enum WordPrintOut {
	WordPrintOutApplication = 'capp',
	WordPrintOutDocument = 'docu',
	WordPrintOutWindow = 'cwin'
};
typedef enum WordPrintOut WordPrintOut;

enum WordRemoveNumbers {
	WordRemoveNumbersDocument = 'docu',
	WordRemoveNumbersListFormat = 'w123',
	WordRemoveNumbersWordList = 'w236'
};
typedef enum WordRemoveNumbers WordRemoveNumbers;

enum WordSmallScroll {
	WordSmallScrollWindow = 'cwin',
	WordSmallScrollPane = 'w120'
};
typedef enum WordSmallScroll WordSmallScroll;

enum WordUpdatePageNumbers {
	WordUpdatePageNumbersTableOfFigures = 'w184',
	WordUpdatePageNumbersTableOfContents = 'w198'
};
typedef enum WordUpdatePageNumbers WordUpdatePageNumbers;

enum WordUpdate {
	WordUpdateLinkFormat = 'w167',
	WordUpdateFieldOptions = 'w298',
	WordUpdateTableOfFigures = 'w184',
	WordUpdateTableOfContents = 'w198',
	WordUpdateTableOfAuthorities = 'w200',
	WordUpdateDialog = 'w202',
	WordUpdateIndex = 'w215'
};
typedef enum WordUpdate WordUpdate;

enum WordWdThemeColorIndex {
	WordWdThemeColorIndexNotThemeColor = '\002\312\377\377',
	WordWdThemeColorIndexThemeColorMainDark1 = '\002\313\000\000',
	WordWdThemeColorIndexThemeColorMainLight1 = '\002\313\000\001',
	WordWdThemeColorIndexThemeColorMainDark2 = '\002\313\000\002',
	WordWdThemeColorIndexThemeColorMainLight2 = '\002\313\000\003',
	WordWdThemeColorIndexThemeColorAccent1 = '\002\313\000\004',
	WordWdThemeColorIndexThemeColorAccent2 = '\002\313\000\005',
	WordWdThemeColorIndexThemeColorAccent3 = '\002\313\000\006',
	WordWdThemeColorIndexThemeColorAccent4 = '\002\313\000\007',
	WordWdThemeColorIndexThemeColorAccent5 = '\002\313\000\010',
	WordWdThemeColorIndexThemeColorAccent6 = '\002\313\000\011',
	WordWdThemeColorIndexThemeColorHyperlink = '\002\313\000\012',
	WordWdThemeColorIndexThemeColorHyperlinkFollowed = '\002\313\000\013',
	WordWdThemeColorIndexThemeColorBackground1 = '\002\313\000\014',
	WordWdThemeColorIndexThemeColorText1 = '\002\313\000\015',
	WordWdThemeColorIndexThemeColorBackground2 = '\002\313\000\016',
	WordWdThemeColorIndexThemeColorText2 = '\002\313\000\017'
};
typedef enum WordWdThemeColorIndex WordWdThemeColorIndex;

enum WordAutomaticLength {
	WordAutomaticLengthShape = 'pShp',
	WordAutomaticLengthCalloutFormat = 'w275'
};
typedef enum WordAutomaticLength WordAutomaticLength;

enum WordCustomDrop {
	WordCustomDropCallout = 'cD00',
	WordCustomDropCalloutFormat = 'w275'
};
typedef enum WordCustomDrop WordCustomDrop;

enum WordCustomLength {
	WordCustomLengthCallout = 'cD00',
	WordCustomLengthCalloutFormat = 'w275'
};
typedef enum WordCustomLength WordCustomLength;

enum WordPresetDrop {
	WordPresetDropCallout = 'cD00',
	WordPresetDropCalloutFormat = 'w275'
};
typedef enum WordPresetDrop WordPresetDrop;

enum WordOneColorGradient {
	WordOneColorGradientShape = 'pShp',
	WordOneColorGradientFillFormat = 'w278'
};
typedef enum WordOneColorGradient WordOneColorGradient;

enum WordPatterned {
	WordPatternedShape = 'pShp',
	WordPatternedFillFormat = 'w278'
};
typedef enum WordPatterned WordPatterned;

enum WordPresetGradient {
	WordPresetGradientShape = 'pShp',
	WordPresetGradientFillFormat = 'w278'
};
typedef enum WordPresetGradient WordPresetGradient;

enum WordPresetTextured {
	WordPresetTexturedShape = 'pShp',
	WordPresetTexturedFillFormat = 'w278'
};
typedef enum WordPresetTextured WordPresetTextured;

enum WordSolid {
	WordSolidShape = 'pShp',
	WordSolidFillFormat = 'w278'
};
typedef enum WordSolid WordSolid;

enum WordTwoColorGradient {
	WordTwoColorGradientShape = 'pShp',
	WordTwoColorGradientFillFormat = 'w278'
};
typedef enum WordTwoColorGradient WordTwoColorGradient;

enum WordUserPicture {
	WordUserPictureShape = 'pShp',
	WordUserPictureFillFormat = 'w278'
};
typedef enum WordUserPicture WordUserPicture;

enum WordUserTextured {
	WordUserTexturedShape = 'pShp',
	WordUserTexturedFillFormat = 'w278'
};
typedef enum WordUserTextured WordUserTextured;

enum WordResetRotation {
	WordResetRotationShape = 'pShp',
	WordResetRotationThreeDFormat = 'w286'
};
typedef enum WordResetRotation WordResetRotation;

enum WordSaveAsPicture {
	WordSaveAsPictureShape = 'pShp',
	WordSaveAsPictureInlineShape = 'w257'
};
typedef enum WordSaveAsPicture WordSaveAsPicture;

enum WordCloseUp {
	WordCloseUpParagraph = 'cpar',
	WordCloseUpParagraphFormat = 'w136'
};
typedef enum WordCloseUp WordCloseUp;

enum WordIndentCharWidth {
	WordIndentCharWidthParagraph = 'cpar',
	WordIndentCharWidthParagraphFormat = 'w136'
};
typedef enum WordIndentCharWidth WordIndentCharWidth;

enum WordIndentFirstLineCharWidth {
	WordIndentFirstLineCharWidthParagraph = 'cpar',
	WordIndentFirstLineCharWidthParagraphFormat = 'w136'
};
typedef enum WordIndentFirstLineCharWidth WordIndentFirstLineCharWidth;

enum WordOpenOrCloseUp {
	WordOpenOrCloseUpParagraph = 'cpar',
	WordOpenOrCloseUpParagraphFormat = 'w136'
};
typedef enum WordOpenOrCloseUp WordOpenOrCloseUp;

enum WordOpenUp {
	WordOpenUpParagraph = 'cpar',
	WordOpenUpParagraphFormat = 'w136'
};
typedef enum WordOpenUp WordOpenUp;

enum WordSpace15 {
	WordSpace15Paragraph = 'cpar',
	WordSpace15ParagraphFormat = 'w136'
};
typedef enum WordSpace15 WordSpace15;

enum WordSpace1 {
	WordSpace1Paragraph = 'cpar',
	WordSpace1ParagraphFormat = 'w136'
};
typedef enum WordSpace1 WordSpace1;

enum WordSpace2 {
	WordSpace2Paragraph = 'cpar',
	WordSpace2ParagraphFormat = 'w136'
};
typedef enum WordSpace2 WordSpace2;

enum WordTabHangingIndent {
	WordTabHangingIndentParagraph = 'cpar',
	WordTabHangingIndentParagraphFormat = 'w136'
};
typedef enum WordTabHangingIndent WordTabHangingIndent;

enum WordTabIndent {
	WordTabIndentParagraph = 'cpar',
	WordTabIndentParagraphFormat = 'w136'
};
typedef enum WordTabIndent WordTabIndent;

enum WordSetLeftIndent {
	WordSetLeftIndentRow = 'crow',
	WordSetLeftIndentRowOptions = 'WrOp'
};
typedef enum WordSetLeftIndent WordSetLeftIndent;

enum WordConvertRowToText {
	WordConvertRowToTextRow = 'crow',
	WordConvertRowToTextRowOptions = 'WrOp'
};
typedef enum WordConvertRowToText WordConvertRowToText;

enum WordAutoFit {
	WordAutoFitColumn = 'ccol',
	WordAutoFitColumnOptions = 'WcOp'
};
typedef enum WordAutoFit WordAutoFit;

enum WordTableSort {
	WordTableSortTable = 'ctbl',
	WordTableSortColumn = 'ccol'
};
typedef enum WordTableSort WordTableSort;

enum WordSetTableItemHeight {
	WordSetTableItemHeightRow = 'crow',
	WordSetTableItemHeightCell = 'ccel',
	WordSetTableItemHeightRowOptions = 'WrOp'
};
typedef enum WordSetTableItemHeight WordSetTableItemHeight;

enum WordSetTableItemWidth {
	WordSetTableItemWidthColumn = 'ccol',
	WordSetTableItemWidthCell = 'ccel',
	WordSetTableItemWidthColumnOptions = 'WcOp'
};
typedef enum WordSetTableItemWidth WordSetTableItemWidth;

enum WordSortAscending {
	WordSortAscendingTable = 'ctbl',
	WordSortAscendingColumn = 'ccol'
};
typedef enum WordSortAscending WordSortAscending;

enum WordSortDescending {
	WordSortDescendingTable = 'ctbl',
	WordSortDescendingColumn = 'ccol'
};
typedef enum WordSortDescending WordSortDescending;

@protocol WordGenericMethods

- (void) closeSaving:(WordSaveOptions)saving savingIn:(NSURL *)savingIn;  // Close a document.
- (void) saveIn:(NSURL *)in_ as:(id)as;  // Save a document.
- (void) printWithProperties:(NSDictionary *)withProperties printDialog:(BOOL)printDialog;  // Print a document.
- (void) delete;  // Delete an object.
- (void) duplicateTo:(SBObject *)to withProperties:(NSDictionary *)withProperties;  // Copy an object.
- (void) moveTo:(SBObject *)to;  // Move an object to a new location.
- (WordMathObject *) addMathArgumentBeforeArg:(NSInteger)beforeArg toFunction:(WordMathFunction *)toFunction toMatrixRow:(WordMathMatrixRow *)toMatrixRow toMatrixColumn:(WordMathMatrixColumn *)toMatrixColumn;  // Inserts an argument into an equation with variable number of arguments, i.e. math delimiter and math equation array objects, and returns a math object object. You must specify only one function, row, or argument.
- (void) autosaveDocument:(WordDocument *)document;  // Causes an autosave to occur for the given document or all documents if no document is specified.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Word can check out a specified document from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified document from a server to a local computer for editing. Returns a String that represents the local path and file name of the document checked out.
- (void) duplicatePage;  // Duplicates the page of the current selection and moves the selection to the new page. This command only works when in Publishing Layout View.
- (void) insertPage;  // Insert a page following the page of the current selection.
- (void) removePage;  // Removes the page of the selection and moves the selection to the following page. When removing the last page, the selection is moved to the previous page. This command only works when in Publishing Layout View.
- (void) threeWayMergeLocalDocument:(WordDocument *)localDocument serverDocument:(WordDocument *)serverDocument baseDocument:(WordDocument *)baseDocument favorSource:(BOOL)favorSource;
- (void) toggleRibbon;  // Toggles whether or not the ribbon is shown in this window.
- (void) toggleShowCodes;

@end



/*
 * Standard Suite
 */

// The application's top-level scripting object.
@interface WordApplication : SBApplication

- (SBElementArray<WordDocument *> *) documents;
- (SBElementArray<WordWindow *> *) windows;

@property (copy, readonly) NSString *name;  // The name of the application.
@property (readonly) BOOL frontmost;  // Is this the active application?
@property (copy, readonly) NSString *version;  // The version number of the application.

- (void) print:(id)x withProperties:(NSDictionary *)withProperties printDialog:(BOOL)printDialog;  // Print a document.
- (void) quitSaving:(WordSaveOptions)saving;  // Quit the application.
- (BOOL) exists:(id)x;  // Verify that an object exists.
- (id) open:(id)x fileName:(NSString *)fileName confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles repair:(BOOL)repair showingRepairs:(BOOL)showingRepairs passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate fileConverter:(WordWdOpenFormat)fileConverter;  // Open a document.
- (id) doScript:(NSString *)x;  // Execute a script
- (void) reset:(id)x;  // Resets a built-in command bar or command bar control to its default configuration.
- (void) WordHelpHelpType:(WordWdHelpType)helpType;  // Displays on-line Help information.
- (void) accept:(id)x;  // Accepts the specified tracked change. The revision marks are removed, and the change is incorporated into the document.
- (void) activateObject:(id)x;  // Activate this object.
- (WordAddIn *) addAddinFileName:(NSString *)fileName install:(BOOL)install;  // Returns an add in object that represents an add-in added to the list of available add-ins.
- (WordCoauthLock *) addCoauthLockToRange:(WordTextRange *)toRange toParagraph:(WordParagraph *)toParagraph toColumn:(WordColumn *)toColumn toCell:(WordCell *)toCell toRow:(WordRow *)toRow toTable:(WordTable *)toTable toSelection:(WordSelectionObject *)toSelection lockType:(WordWdLockType)lockType inRange:(WordSelectionObject *)inRange inCoauthoring:(WordCoauthoring *)inCoauthoring inCoauthor:(WordCoauthor *)inCoauthor;  // Add a co-authoring lock.
- (void) arrangeWindowsArrangeStyle:(WordWdArrangeStyle)arrangeStyle;  // Arrange windows on the screen.
- (NSInteger) buildKeyCodeKey1:(WordWdKey)key1 key2:(WordWdKey)key2 key3:(WordWdKey)key3 key4:(WordWdKey)key4;  // Returns a unique number for the specified key combination.
- (WordWdContinue) canContinuePreviousList:(id)x listTemplate:(WordListTemplate *)listTemplate;  // Returns whether the formatting from the previous list can be continued.
- (double) centimetersToPointsCentimeters:(double)centimeters;  // Converts a measurement from centimeters to points.
- (void) changeFileOpenDirectoryPath:(NSString *)path;  // Sets the folder in which Microsoft Word searches for documents.
- (BOOL) checkGrammar:(id)x textToCheck:(NSString *)textToCheck;  // Checks a string for grammatical errors. Returns a boolean to indicate whether the string contains grammatical errors. True if the string contains no errors.
- (BOOL) checkSpelling:(id)x textToCheck:(NSString *)textToCheck customDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Checks a string for spelling errors. Returns true if the string has no spelling errors.
- (NSString *) cleanStringItemToCheck:(NSString *)itemToCheck;  // Removes nonprinting characters character codes 1 Ôøê 29 and special Microsoft Word characters from the specified string or changes them to spaces character code 32. Returns the result as a string.
- (void) clear:(id)x;  // Removes the object.
- (void) clearFormatting:(id)x;  // Removes text and paragraph formatting from a selection or from the formatting specified in a find or replace operation.
- (void) convertNumbersToText:(id)x numberType:(WordWdNumberType)numberType;  // Changes the list numbers and listnum fields in the document object to text.
- (void) copyObject:(id)x NS_RETURNS_NOT_RETAINED;  // Copies the content of the object to the clipboard.
- (NSInteger) countNumberedItems:(id)x numberType:(WordWdNumberType)numberType level:(NSInteger)level;  // Returns the number of bulleted or numbered items and listnum fields in the document object.
- (WordDocument *) createNewDocumentAttachedTemplate:(WordTemplate *)attachedTemplate newTemplate:(BOOL)newTemplate newDocumentType:(WordWdNewDocumentType)newDocumentType;  // Create a new document
- (void) createNewEquationFromRange:(WordTextRange *)fromRange inDocument:(WordDocument *)inDocument inRange:(WordSelectionObject *)inRange inSelection:(WordSelectionObject *)inSelection;  // Creates an equation, from the text equation contained within the specified range, and returns a text range object that contains the new equation.
- (void) createNewFieldTextRange:(WordTextRange *)textRange fieldType:(WordWdFieldType)fieldType fieldText:(NSString *)fieldText preserveFormatting:(BOOL)preserveFormatting;  // Create a new field
- (void) cutObject:(id)x;  // Removes the object from the document and places it on the clipboard.
- (BOOL) doWordRepeatTimes:(NSInteger)times;  // Repeats the most recent editing action one or more times. Returns true if the commands were repeated successfully.
- (WordKeyBinding *) findKeyKeyCode:(NSInteger)keyCode key_code_2:(WordWdKey)key_code_2;  // Returns a key binding object that represents the specified key combination
- (WordBorder *) getBorder:(id)x whichBorder:(WordWdBorderType)whichBorder;  // Returns the specified border object.
- (NSString *) getDefaultFilePathFilePathType:(WordWdDefaultFilePath)filePathType;  // Returns the default folders for items such as documents, templates, and graphics.
- (WordDialog *) getDialog:(WordWdWordDialog)x;  // Returns a dialog object for the specified dialog.
- (NSString *) getInternationalInformation:(WordWdInternationalIndex)x;  // Get the specified international information
- (NSArray<SBObject *> *) getKeysBoundToKeyCategory:(WordWdKeyCategory)keyCategory command:(NSString *)command;  // Returns a list key binding objects that represents all the key combinations assigned to the specified item.
- (WordListGallery *) getListGallery:(WordWdListGalleryType)x;  // Returns the specified list gallery object object.
- (NSDictionary *) getSpellingSuggestionsItemToCheck:(NSString *)itemToCheck customDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary suggestionMode:(WordWdSpellingWordType)suggestionMode customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Returns a record that specifies the error kind and a list of words.  The AEKeyword for the error kind is type and AEKeyword for the list is list.
- (WordSynonymInfo *) getSynonymInfoObjectItemToCheck:(NSString *)itemToCheck languageID:(WordWdLanguageID)languageID;  // Returns a synonym info object that contains the information from the thesaurus on the synonyms, antonyms, or related words and expressions for the specified word or phrase.
- (WordWebPageFont *) getWebpageFont:(WordMsoCharacterSet)x;  // Return the specified web page font object for a given character set.
- (double) inchesToPointsInches:(double)inches;  // Converts a measurement from inches to points.
- (void) insertText:(NSString *)text at:(SBObject *)at;  // Insert the string at the specified location.
- (void) insertAutoTextAt:(WordTextRange *)at;  // Attempts to match the text in the specified range or the text surrounding the range with an existing auto text entry name.  If any such match is found, the auto text entry is inserted.  If no match an error occurs.
- (void) insertBreakAt:(WordTextRange *)at breakType:(WordWdBreakType)breakType;  // Inserts a break in the specified place of the specified kind.
- (void) insertCaptionAt:(WordTextRange *)at captionLabel:(WordWdCaptionLabelID)captionLabel title:(NSString *)title captionPosition:(WordWdCaptionPosition)captionPosition;  // Inserts a caption immediately preceding or following the specified range.
- (void) insertCrossReferenceAt:(WordTextRange *)at referenceType:(WordWdReferenceType)referenceType referenceKind:(WordWdReferenceKind)referenceKind referenceItem:(NSString *)referenceItem insertAsHyperlink:(BOOL)insertAsHyperlink includePosition:(BOOL)includePosition;  // Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined.
- (void) insertDatabaseAt:(WordTextRange *)at format:(WordWdTableFormat)format style:(NSInteger)style linkToSource:(BOOL)linkToSource connection:(NSString *)connection SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1 passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate dataSource:(NSString *)dataSource from:(NSInteger)from to:(NSInteger)to includeFields:(BOOL)includeFields;  // Retrieves data from a data source and inserts the data as a table in place of the specified range.
- (void) insertDateTimeAt:(WordTextRange *)at dateTimeFormat:(NSString *)dateTimeFormat insertAsField:(BOOL)insertAsField;  // Insert the correct date or time, or both, either as text or as a time field at the specified location.
- (void) insertFileAt:(WordTextRange *)at fileName:(NSString *)fileName fileRange:(NSString *)fileRange confirmConversions:(BOOL)confirmConversions link:(BOOL)link;
- (void) insertParagraphAt:(WordTextRange *)at;  // Replaces the specified range with a new paragraph.  If you do not want to replace the range, use the collapse range method before using this method.
- (void) insertSymbolAt:(WordTextRange *)at characterNumber:(NSInteger)characterNumber font:(NSString *)font unicode:(BOOL)unicode bias:(WordWdFontBias)bias;  // Inserts a symbol in place of the specified range.  If you do not want to replace the range, use the collapse range method before using this method.
- (NSString *) keyStringKeyCode:(NSInteger)keyCode key_code_2:(WordWdKey)key_code_2;  // Returns the key combination string for the specified keys 
- (void) largeScroll:(id)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls a window by the specified number of screens. This method is equivalent to clicking just before or just after the scroll boxes on the horizontal and vertical scroll bars.
- (double) linesToPointsLines:(double)lines;  // Converts a measurement from lines to points. 1 line = 12 points.
- (void) listCommandsListAllCommands:(BOOL)listAllCommands;  // Creates a new document and then inserts a table of Microsoft Word commands along with their associated shortcut keys and menu assignments.
- (double) millimetersToPointsMillimeters:(double)millimeters;  // Converts a measurement from millimeters to points.
- (void) organizerCopySource:(NSString *)source destination:(NSString *)destination name:(NSString *)name organizerObjectType:(WordWdOrganizerObject)organizerObjectType;  // Copies the specified autotext entry, toolbar, style, or macro project item from the source document or template to the destination document or template.
- (void) organizerDeleteSource:(NSString *)source name:(NSString *)name organizerObjectType:(WordWdOrganizerObject)organizerObjectType;  // Deletes the specified style, autotext entry, toolbar, or macro project item from a document or template.
- (void) organizerRenameSource:(NSString *)source name:(NSString *)name newName:(NSString *)newName organizerObjectType:(WordWdOrganizerObject)organizerObjectType;  // Renames the specified style, autotext entry, toolbar, or macro project item in a document or template.
- (void) pageScroll:(id)x down:(NSInteger)down up:(NSInteger)up;  // Scrolls through the window page by page.
- (double) picasToPointsPicas:(double)picas;  // Converts a measurement from picas to points.
- (double) pointsToCentimetersPoints:(double)points;  // Converts a measurement from points to centimeters.
- (double) pointsToInchesPoints:(double)points;  // Converts a measurement from points to inches.
- (double) pointsToLinesPoints:(double)points;  // Converts a measurement from points to lines. 1 line = 12 points.
- (double) pointsToMillimetersPoints:(double)points;  // Converts a measurement from points to millimeters.
- (double) pointsToPicasPoints:(double)points;  // Converts a measurement from points to picas.
- (void) printOut:(id)x append:(BOOL)append printOutRange:(WordWdPrintOutRange)printOutRange outputFileName:(NSString *)outputFileName pageFrom:(NSInteger)pageFrom pageTo:(NSInteger)pageTo printOutItem:(WordWdPrintOutItem)printOutItem printCopies:(NSInteger)printCopies printOutPageType:(WordWdPrintOutPages)printOutPageType printToFile:(BOOL)printToFile collate:(BOOL)collate fileName:(NSString *)fileName manualDuplexPrint:(BOOL)manualDuplexPrint;  // Prints out all or part of the specified or active document. This command has been deprecated; use the Print command in the Standard Suite.
- (void) reject:(id)x;  // Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.
- (void) removeNumbers:(id)x numberType:(WordWdNumberType)numberType;  // Removes numbers or bullets from the document
- (void) remveEmphemeralLocksInRange:(WordSelectionObject *)inRange inCoauthoring:(WordCoauthoring *)inCoauthoring inCoauthor:(WordCoauthor *)inCoauthor;
- (void) resetContinuationNotice:(id)x;  // Resets the footnote or endnote continuation notice to the default notice. The default notice is blank.
- (void) resetContinuationSeparator:(id)x;  // Resets the footnote or endnote continuation separator to the default separator. The default separator is a long horizontal line that separates document text from notes continued from the previous page.
- (void) resetIgnoreAll;  // Clears the list of words that were previously ignored during a spelling check. After you run this method, previously ignored words are checked along with all the other words.
- (void) resetSeparator:(id)x;  // Resets the footnote or endnote separator to the default separator. The default separator is a short horizontal line that separates document text from notes.
- (WordLanguage *) retrieveLanguage:(WordWdLanguageID)x;  // Returns the language object for the specified language.
- (void) runVBMacroMacroName:(NSString *)macroName;  // Runs a Visual Basic macro.
- (void) screenRefresh;  // Updates the display on the monitor with the current information in the video memory buffer. You can use this method after using the screen updating property to disable screen updates.
- (void) setDefaultFilePathFilePathType:(WordWdDefaultFilePath)filePathType path:(NSString *)path;  // Sets the default folders for items such as documents, templates, and graphics.
- (void) smallScroll:(id)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls a window by the specified number of lines. This method is equivalent to clicking the scroll arrows on the horizontal and vertical scroll bars.
- (void) substituteFontUnavailableFont:(NSString *)unavailableFont substituteFont:(NSString *)substituteFont;  // Sets font-mapping options, which are reflected in the font substitution dialog box
- (void) unlink:(id)x;
- (void) unloadAddinsRemoveFromList:(BOOL)removeFromList;  // Unloads all loaded add-ins and, depending on the value of the remove from list argument, removes them from the list of add-ins.
- (void) update:(id)x;  // Updates the values shown in a built-in Microsoft Word dialog box, updates the specified link, or updates the entries shown in specified index, table of authorities, table of figures or table of contents.
- (void) updatePageNumbers:(id)x;  // Updates the page numbers for items in the specified table of contents or table of figures.
- (void) updateSource:(id)x;
- (void) automaticLength:(id)x;  // Specifies that the first segment of the callout line the segment attached to the text callout box be scaled automatically when the callout is moved. Applies only to callouts whose lines consist of more than one segment.
- (void) customDrop:(id)x drop:(double)drop;  // Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
- (void) customLength:(id)x length:(double)length;  // Specifies that the first segment of the callout, line the segment attached to the text callout box, retain a fixed length whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment.
- (void) oneColorGradient:(id)x gradientStyle:(WordMsoGradientStyle)gradientStyle gradientVariant:(NSInteger)gradientVariant gradientDegree:(double)gradientDegree;  // Sets the specified fill to a one-color gradient.
- (void) patterned:(id)x pattern:(WordMsoPatternType)pattern;  // Sets the specified fill to a pattern.
- (void) presetDrop:(id)x DropType:(WordMsoCalloutDropType)DropType;  // Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
- (void) presetGradient:(id)x style:(WordMsoGradientStyle)style gradientVariant:(NSInteger)gradientVariant presetGradientType:(WordMsoPresetGradientType)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetTextured:(id)x presetTexture:(WordMsoPresetTexture)presetTexture;  // Sets the specified fill to a preset texture
- (void) resetRotation:(id)x;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- (void) saveAsPicture:(id)x pictureType:(WordMsoPictureType)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) solid:(id)x;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) twoColorGradient:(id)x gradientStyle:(WordMsoGradientStyle)gradientStyle gradientVariant:(NSInteger)gradientVariant;  // Sets the specified fill to a two-color gradient.
- (void) userPicture:(id)x pictureFile:(NSString *)pictureFile;  // Fills the specified shape with one large image. If you want to fill the shape with small tiles of an image, use the user textured method.
- (void) userTextured:(id)x textureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the user picture method.
- (void) closeUp:(id)x;  // Removes any spacing before the specified paragraphs.
- (WordBorder *) getBorder:(id)x whichBorder:(WordWdBorderType)whichBorder;  // Returns the specified border object.
- (void) indentCharWidth:(id)x count:(NSInteger)count;  // Indents one or more paragraphs by a specified number of characters.
- (void) indentFirstLineCharWidth:(id)x count:(NSInteger)count;  // Indents the first line of one or more paragraphs by a specified number of characters.
- (void) openOrCloseUp:(id)x;  // If spacing before the specified paragraphs is zero, this method sets spacing to 12 points. If spacing before the paragraphs is greater than zero, this method sets spacing to zero.
- (void) openUp:(id)x;  // Sets spacing before the specified paragraphs to 12 points.
- (void) reset:(id)x;  // Removes paragraph formatting that differs from the underlying style. For example, if you manually right align a paragraph and the underlying style has a different alignment, the reset method changes the alignment to match the style formatting.
- (void) space1:(id)x;  // Single-spaces the specified paragraphs. The exact spacing is determined by the font size of the largest characters in each paragraph.
- (void) space15:(id)x;  // Formats the specified paragraphs with 1.5-line spacing. The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.
- (void) space2:(id)x;  // Double-spaces the specified paragraphs. The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.
- (void) tabHangingIndent:(id)x count:(NSInteger)count;  // Sets a hanging indent to a specified number of tab stops. Can be used to remove tab stops from a hanging indent if the value of the count argument is a negative number.
- (void) tabIndent:(id)x count:(NSInteger)count;  // Sets the left indent for the specified paragraphs to a specified number of tab stops. Can also be used to remove the indent if the value of the count argument is a negative number.
- (void) autoFit:(id)x;  // Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
- (WordTextRange *) convertRowToText:(id)x separator:(WordWdTableFieldSeparator)separator nestedTables:(BOOL)nestedTables;  // Converts a row to text and returns a text range object that represents the delimited text.
- (WordBorder *) getBorder:(id)x whichBorder:(WordWdBorderType)whichBorder;  // Returns the specified border object.
- (void) setLeftIndent:(id)x leftIndent:(double)leftIndent rulerStyle:(WordWdRulerStyle)rulerStyle;  // Sets the indentation for a row or rows in a table.
- (void) setTableItemHeight:(id)x rowHeight:(NSInteger)rowHeight heightRule:(WordWdRowHeightRule)heightRule;  // Sets the height of table rows.
- (void) setTableItemWidth:(id)x columnWidth:(double)columnWidth rulerStyle:(WordWdRulerStyle)rulerStyle;  // Sets the width of columns or cells in a table
- (void) sortAscending:(id)x;  // Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) sortDescending:(id)x;  // Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) tableSort:(id)x excludeHeader:(BOOL)excludeHeader fieldNumber:(NSInteger)fieldNumber sortFieldType:(WordWdSortFieldType)sortFieldType sortOrder:(WordWdSortOrder)sortOrder fieldNumberTwo:(NSInteger)fieldNumberTwo sortFieldTypeTwo:(WordWdSortFieldType)sortFieldTypeTwo sortOrderTwo:(WordWdSortOrder)sortOrderTwo fieldNumberThree:(NSInteger)fieldNumberThree sortFieldTypeThree:(WordWdSortFieldType)sortFieldTypeThree sortOrderThree:(WordWdSortOrder)sortOrderThree sortColumn:(BOOL)sortColumn separator:(WordWdSortSeparator)separator caseSensitive:(BOOL)caseSensitive languageID:(WordWdLanguageID)languageID;  // Sort a table object.  For the column object only the first field is used

@end

// A document.
@interface WordDocument : SBObject <WordGenericMethods>

@property (copy, readonly) NSString *name;  // Its name.
@property (readonly) BOOL modified;  // Has it been modified since the last save?
@property (copy, readonly) NSURL *file;  // Its location on disk, if it has one.

- (void) acceptAllRevisions;  // Accepts all tracked changes in the document.
- (void) acceptAllShownRevisions;  // Accepts all shown tracked changes.
- (void) applyDocumentThemeFileName:(NSString *)fileName;  // Applies an Office theme to the document.
- (void) autoFormat;  // Automatically formats a document
- (BOOL) canCheckIn;  // Returns True if Word can check in a specified document to a server.
- (void) changeDefaultTableStyleStyle:(WordWdBuiltinStyle)style changeInTemplate:(BOOL)changeInTemplate;  // Sets the default table style.
- (void) checkConsistency;  // Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a document from a local computer to a server, and sets the local document to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(WordWdCheckInVersionType)versionType;  // Returns a document from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) closePrintPreview;  // Switches the specified document from print preview to the previous view. If the specified document isn't in print preview, an error occurs.
- (void) comparePath:(NSString *)path authorName:(NSString *)authorName target:(WordWdCompareTarget)target detectFormatChanges:(BOOL)detectFormatChanges ignoreAllComparisonWarnings:(BOOL)ignoreAllComparisonWarnings addToRecentFiles:(BOOL)addToRecentFiles;  // Compares this document with another.
- (NSInteger) computeStatisticsStatistic:(WordWdStatistic)statistic includeFootnotesAndEndnotes:(BOOL)includeFootnotesAndEndnotes;  // Computes a set of readability statistics for this document.  This must be called before accessing the readability statistics for this document.
- (void) copyStylesFromTemplateTemplate:(NSString *)template_ NS_RETURNS_NOT_RETAINED;  // Copies styles from the specified template to a document.
- (WordLetterContent *) createLetterContentDateFormat:(NSString *)dateFormat includeHeaderFooter:(BOOL)includeHeaderFooter pageDesign:(NSString *)pageDesign letterStyle:(WordWdLetterStyle)letterStyle letterhead:(BOOL)letterhead letterheadLocation:(WordWdLetterheadLocation)letterheadLocation letterheadSize:(double)letterheadSize recipientName:(NSString *)recipientName recipientAddress:(NSString *)recipientAddress salutation:(NSString *)salutation salutationType:(WordWdSalutationType)salutationType recipientReference:(NSString *)recipientReference mailingInstructions:(NSString *)mailingInstructions attentionLine:(NSString *)attentionLine subject:(NSString *)subject ccList:(NSString *)ccList returnAddress:(NSString *)returnAddress senderName:(NSString *)senderName closing:(NSString *)closing senderCompany:(NSString *)senderCompany senderJobTitle:(NSString *)senderJobTitle senderInitials:(NSString *)senderInitials enclosureCount:(NSInteger)enclosureCount;  // Create a new letter content object for use with the letter wizard.
- (WordTextRange *) createRangeStart:(NSInteger)start end:(NSInteger)end;  // Returns a text range object by using the specified starting and ending character positions.
- (void) dataForm;  // Displays the data form dialog box, in which you can modify data records. You can use this method with a mail merge main document, a mail merge data source, or any document that contains data delimited by table cells or separator characters
- (void) deleteAllComments;  // Deletes all comments.
- (void) deleteAllShownComments;  // Deletes all shown comments.
- (void) fitToPages;  // Decreases the font size of text just enough so that the document will fit on one fewer pages. An error occurs if Word is unable to reduce the page count by one.
- (void) followHyperlinkAddress:(NSString *)address subAddress:(NSString *)subAddress newWindow:(BOOL)newWindow addHistory:(BOOL)addHistory extraInfo:(NSString *)extraInfo;  // This method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application. If the hyperlink uses the file protocol, this method opens the document instead of downloading it.
- (NSString *) getActiveWritingStyleLanguageID:(WordWdLanguageID)languageID;  // Returns the writing style for a specified language.
- (NSArray<NSString *> *) getCrossReferenceItemsReferenceType:(WordWdReferenceType)referenceType;  // Returns an list of items that can be cross-referenced based on the specified cross-reference type.
- (BOOL) getDocumentCompatibilityCompatibilityItem:(WordWdCompatibility)compatibilityItem;  // Returns the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.
- (WordStoryRange *) getStoryRangeStoryType:(WordWdStoryType)storyType;  // Returns a range object that represents the story specified by the story type argument.
- (void) makeCompatibilityDefault;  // Uses the correct settings of the document compatibility options set by the set compatibility options method as the default for new documents.
- (void) manualHyphenation;  // Initiates manual hyphenation of a document, one line at a time. The user is prompted to accept or decline suggested hyphenations.
- (WordField *) markEntryForTableOfContentsRange:(WordTextRange *)range entry:(NSString *)entry tableID:(NSString *)tableID level:(NSInteger)level;  // Inserts a table of contents entry field after the specified range. The method returns a field object representing the new field.
- (WordField *) markEntryForTableOfFiguresRange:(WordTextRange *)range entry:(NSString *)entry tableID:(NSString *)tableID level:(NSInteger)level;  // Inserts a table of figures entry field after the specified range. The method returns a field object representing the new field.
- (WordField *) markForIndexRange:(WordTextRange *)range entry:(NSString *)entry crossReference:(NSString *)crossReference bookmarkName:(NSString *)bookmarkName;  // Inserts an index entry field after the specified range. The method returns a field object representing the new field.
- (void) mergeFileName:(NSString *)fileName;  // Merges the changes marked with revision marks from one document to another.
- (void) mergeSubdocumentsFirstSubdocument:(WordSubdocument *)firstSubdocument lastSubdocument:(WordSubdocument *)lastSubdocument;  // Merges the specified subdocuments of a master document into a single subdocument.
- (void) printPreview;  // Switches the view to print preview.
- (void) protectProtectionType:(WordWdProtectionType)protectionType noReset:(BOOL)noReset password:(NSString *)password useIrm:(BOOL)useIrm enforceStyleLocks:(BOOL)enforceStyleLocks;  // Helps to protect the specified document from changes. When a document is protected, users can make only limited changes, such as adding annotations, making revisions, or completing a form.
- (BOOL) redoTimes:(NSInteger)times;  // Redoes the last action that was undone. It reverses the undo method. Returns true if the actions were redone successfully.
- (void) rejectAllRevisions;  // Rejects all tracked changes in the document.
- (void) rejectAllShownRevisions;  // Rejects all shown tracked changes.
- (void) reload;  // Reloads a cached document by resolving the hyperlink to the document and downloading it.
- (void) repaginate;  // Repaginates the entire document
- (void) runLetterWizardLetterContent:(WordLetterContent *)letterContent wizardMode:(BOOL)wizardMode;  // Runs the Letter Wizard on the document
- (void) saveAsFileName:(NSString *)fileName fileFormat:(WordWdSaveFormat)fileFormat lockComments:(BOOL)lockComments password:(NSString *)password addToRecentFiles:(BOOL)addToRecentFiles writePassword:(NSString *)writePassword readOnlyRecommended:(BOOL)readOnlyRecommended embedTruetypeFonts:(BOOL)embedTruetypeFonts saveNativePictureFormat:(BOOL)saveNativePictureFormat saveFormsData:(BOOL)saveFormsData textEncoding:(NSInteger)textEncoding insertLineBreaks:(BOOL)insertLineBreaks allowSubstitutions:(BOOL)allowSubstitutions lineEndingType:(WordWdLineEndingType)lineEndingType AddBiDiMarks:(BOOL)AddBiDiMarks;  // Saves the document with a new name or format.
- (void) saveVersionComment:(NSString *)comment;  // Saves a version of the document with a comment.
- (void) sendHtmlMail;  // Opens a message window for sending the specified document, formatted as html, through Microsoft Outlook.
- (void) sendMail;  // Opens a message window for sending the specified document through your registered mail program.
- (void) setActiveWritingStyleLanguageID:(WordWdLanguageID)languageID writingStyle:(NSString *)writingStyle;  // Sets the writing style for the specified language.
- (void) setDocumentCompatibilityCompatibilityItem:(WordWdCompatibility)compatibilityItem isCompatible:(BOOL)isCompatible;  // Sets the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.
- (BOOL) undoTimes:(NSInteger)times;  // Undoes the last action or a sequence of actions, which are displayed in the undo list. Returns true if the actions were successfully undone.
- (void) undoClear;  // Clear the list of actions that can be undone.
- (void) unprotectPassword:(NSString *)password;  // Removes protection from the specified document. If the document isn't protected, this method generates an error.
- (void) updateStyles;  // Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.
- (void) upgrade;  // upgrade document
- (void) viewPropertyBrowser;  // Displays the property window for the selected control in the specified document.
- (void) webPagePreview;  // Displays a preview of the document as it would look if saved as a Web page.

@end

// A window.
@interface WordWindow : SBObject <WordGenericMethods>

@property (copy, readonly) NSString *name;  // The title of the window.
- (NSInteger) id;  // The unique identifier of the window.
@property NSInteger index;  // The index of the window, ordered front to back.
@property NSRect bounds;  // The bounding rectangle of the window.
@property (readonly) BOOL closeable;  // Does the window have a close button?
@property (readonly) BOOL miniaturizable;  // Does the window have a minimize button?
@property BOOL miniaturized;  // Is the window minimized right now?
@property (readonly) BOOL resizable;  // Can the window be resized?
@property BOOL visible;  // Is the window visible right now?
@property (readonly) BOOL zoomable;  // Does the window have a zoom button?
@property BOOL zoomed;  // Is the window zoomed right now?
@property (copy, readonly) WordDocument *document;  // The document whose contents are displayed in the window.
@property (readonly) NSInteger entryIndex;  // the number of the window
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?

- (void) toggleShowAllReviewers;  // Toggles whether or not all reviewers are shown in this window.

@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface WordCommandBarControl : SBObject <WordGenericMethods>

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) WordMsoControlType controlType;  // Returns the type of the command bar control.
@property (copy) NSString *descriptionText;  // Returns or sets the description for a command bar control.  The description is not displayed to the user, but it can be useful for documenting the behavior of a control.
@property BOOL enabled;  // Returns or sets if the command bar control is enabled.
@property (readonly) NSInteger entry_index;  // Returns the index number for this command bar control.
@property NSInteger height;  // Returns or sets the height of a command bar control.
@property NSInteger helpContextID;  // Returns or sets the help context ID number for the Help topic attached to the command bar control.
@property (copy) NSString *helpFile;  // Returns or sets the file name for the help topic attached to the command bar.  To use this property, you must also set the help context ID property.
- (NSInteger) id;  // Returns the id for a built-in command bar control.
@property (readonly) NSInteger leftPosition;  // Returns the left position of the command bar control.
@property (copy) NSString *name;  // Returns or sets the caption text for a command bar control.
@property (copy) NSString *parameter;  // Returns or sets a string that is used to execute a command.
@property NSInteger priority;  // Returns or sets the priority of a command bar control. A controls priority determines whether the control can be dropped from a docked command bar if the command bar controls can not fit in a single row.  Valid priority number are 0 through 7.
@property (copy) NSString *tag;  // Returns or sets information about the command bar control, such as data that can be used as an argument in procedures, or information that identifies the control.
@property (copy) NSString *tooltipText;  // Returns or sets the text displayed in a command bar controls tooltip.
@property (readonly) NSInteger top;  // Returns the top position of a command bar control.
@property BOOL visible;  // Returns or sets if the command bar control is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar control.

- (void) execute;  // Runs the procedure or built-in command assigned to the specified command bar control.

@end

// A button control within a command bar.
@interface WordCommandBarButton : WordCommandBarControl

@property (readonly) BOOL buttonFaceIsDefault;  // Returns if the face of a command bar button control is the original built-in face.
@property WordMsoButtonState buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property WordMsoButtonStyle buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface WordCommandBarCombobox : WordCommandBarControl

@property WordMsoComboStyle comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
@property (copy) NSString *comboboxText;  // Returns or sets the text in the display or edit portion of the command bar combobox control.
@property NSInteger dropDownLines;  // Returns or sets the number of lines in a command bar control combobox control.  The combobox control must be a custom control.
@property NSInteger dropDownWidth;  // Returns or sets the width in pixels of the list for the specified command bar combobox control.  An error occurs if you attempt to set this property for a built-in combobox control.
@property NSInteger listIndex;

- (void) addItemToComboboxComboboxItem:(NSString *)comboboxItem entry_index:(NSInteger)entry_index;  // Add a new string to a custom combobox control.
- (void) clearCombobox;  // Clear all of the strings form a custom combobox.
- (NSString *) getComboboxItemEntry_index:(NSInteger)entry_index;  // Return the string at the given index within a combobox.
- (NSInteger) getCountOfComboboxItems;  // Return the number of strings within a combobox.
- (void) removeAnItemFromComboboxEntry_index:(NSInteger)entry_index;  // Remove a string from a custom combobox.
- (void) setComboboxItemEntry_index:(NSInteger)entry_index comboboxItem:(NSString *)comboboxItem;  // Set the string an a given index for a custom combobox.

@end

// A popup menu control within a command bar.
@interface WordCommandBarPopup : WordCommandBarControl

- (SBElementArray<WordCommandBarControl *> *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface WordCommandBar : SBObject <WordGenericMethods>

- (SBElementArray<WordCommandBarControl *> *) commandBarControls;

@property WordMsoBarPosition barPosition;  // Returns or sets the position of the command bar.
@property (readonly) WordMsoBarType barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray<NSAppleEventDescriptor *> *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

@interface WordDocumentProperty : SBObject <WordGenericMethods>

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface WordCustomDocumentProperty : WordDocumentProperty


@end

@interface WordWebPageFont : SBObject <WordGenericMethods>

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft Word Suite
 */

// Represents a single comment.
@interface WordWordComment : SBObject <WordGenericMethods>

@property (copy) NSString *author;
@property (readonly) NSInteger commentIndex;
@property (copy, readonly) WordTextRange *commentText;
@property (copy, readonly) NSDate *dateValue;
@property (copy) NSString *initials;
@property (copy, readonly) WordTextRange *noteReference;
@property (copy, readonly) WordTextRange *scope;
@property (readonly) BOOL showTip;


@end

// Represents a single list format that's been applied to specified paragraphs in a document.
@interface WordWordList : SBObject <WordGenericMethods>

- (SBElementArray<WordParagraph *> *) paragraphs;

@property (readonly) BOOL singleListTemplate;  // Returns if the entire Word list object uses the same list template.
@property (copy, readonly) NSString *styleName;
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the Word list.

- (void) applyListTemplateListTemplate:(WordListTemplate *)listTemplate continuePreviousList:(BOOL)continuePreviousList defaultListBehavior:(WordWdDefaultListBehavior)defaultListBehavior;  // Applies a set of list-formatting characteristics to the specified Word list object.

@end

// Represents application and document options in Word.
@interface WordWordOptions : SBObject <WordGenericMethods>

@property BOOL EnableMisusedWordsDictionary;  // Returns or sets if Microsoft Word checks for misused words when checking the spelling and grammar in a document.
@property BOOL IMEAutomaticControl;  // Returns or sets if Microsoft Word is set to automatically open and close the Japanese Input Method Editor. 
@property BOOL RTFInClipboard;  // Returns or sets if all text copied from Word to the Clipboard retains its character and paragraph formatting.
@property BOOL allowAccentedUppercase;  // Returns or sets if accents are retained when a French language character is changed to uppercase.
@property BOOL allowClickAndTypeMouse;  // Returns or sets if click and type functionality is enabled.
@property BOOL allowDragAndDrop;  // Returns or sets if dragging and dropping can be used to move or copy a selection.
@property BOOL allowFastSave;  // Returns or sets if Word saves only changes to a document. When reopening the document, Word uses the saved changes to reconstruct the document.
@property BOOL animateScreenMovements;  // Returns or sets if Word animates mouse movements, uses animated cursors, and animates actions such as background saving and find and replace operations.
@property BOOL applyEastAsianFontsToAscii;  // Returns or sets if Microsoft Word applies East Asian fonts to Latin text.
@property BOOL autoFormatApplyBulletedLists;  // Returns or sets if characters, such as asterisks, hyphens, and greater-than signs, at the beginning of list paragraphs are replaced with bullets from the Bullets and Numbering dialog box when Word formats a document or range automatically.
@property BOOL autoFormatApplyHeadings;  // Returns or sets if styles are automatically applied to headings when Word formats a document or range automatically.
@property BOOL autoFormatApplyLists;  // Returns or sets if  if styles are automatically applied to lists when Word formats a document or range automatically.
@property BOOL autoFormatApplyOtherParagraphs;  // Returns or sets if styles are automatically applied to paragraphs that aren't headings or list items when Word formats a document or range automatically.
@property BOOL autoFormatAsYouTypeApplyBorders;  // Returns or sets if a series of three or more hyphens -, equal signs =, or underscore characters _ are automatically replaced by a specific border line when the ENTER key is pressed.
@property BOOL autoFormatAsYouTypeApplyBulletedLists;  // Returns or sets if bullet characters, such as asterisks, hyphens, and greater-than signs, are replaced with bullets from the bullets and numbering dialog box as you type.
@property BOOL autoFormatAsYouTypeApplyClosings;  // Returns or sets if Microsoft Word to automatically applies the closing style to letter closings as you type.
@property BOOL autoFormatAsYouTypeApplyDates;  // Returns or sets if Microsoft Word  automatically applies the date style to dates as you type.
@property BOOL autoFormatAsYouTypeApplyFirstIndents;  // Returns or sets if Microsoft Word automatically replaces a space entered at the beginning of a paragraph with a first-line indent.
@property BOOL autoFormatAsYouTypeApplyHeadings;  // Returns or sets if styles are automatically applied to headings as you type.
@property BOOL autoFormatAsYouTypeApplyNumberedLists;  // Returns or sets if paragraphs are automatically formatted as numbered lists with a numbering scheme from the Bullets and Numbering dialog box according to what's typed.
@property BOOL autoFormatAsYouTypeApplyTables;  // Returns or set if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. The plus signs become the column borders, and the hyphens become the column widths.
@property BOOL autoFormatAsYouTypeAutoLetterWizard;  // Returns or sets if Microsoft Word to automatically starts the Letter Wizard when the user enters a letter salutation or closing.
@property BOOL autoFormatAsYouTypeDefineStyles;  // Returns or sets if Word automatically creates new styles based on manual formatting.
@property BOOL autoFormatAsYouTypeDeleteAutoSpaces;  // Returns or sets if Microsoft Word to automatically deletes spaces inserted between Japanese and Latin text as you type.
@property BOOL autoFormatAsYouTypeFormatListItemBeginning;  // Returns or sets if Word repeats character formatting applied to the beginning of a list item to the next list item.
@property BOOL autoFormatAsYouTypeInsertClosings;  // Returns or sets if Microsoft Word to automatically inserts the corresponding memo closing when the user enters a memo heading.
@property BOOL autoFormatAsYouTypeInsertOvers;  // Returns or sets if Word will automatically inset certain Japanese characters for other Japanese characters.
@property BOOL autoFormatAsYouTypeMatchParentheses;  // Returns or sets if Microsoft Word to automatically corrects improperly paired parentheses.
@property BOOL autoFormatAsYouTypeReplaceEastAsianDashes;  // Returns or sets if Microsoft Word to automatically corrects long vowel sounds and dashes.
@property BOOL autoFormatAsYouTypeReplaceFractions;  // Returns or sets if typed fractions are replaced with fractions from the current character set as you type.
@property BOOL autoFormatAsYouTypeReplaceHyperlinks;  // Returns or sets if e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLs, are automatically changed to hyperlinks as you type.
@property BOOL autoFormatAsYouTypeReplaceOrdinals;  // Returns or sets if the ordinal number suffixes st, nd, rd, and th are replaced with the same letters in superscript as you type. For example, 1st is replaced with 1 followed by st formatted as superscript.
@property BOOL autoFormatAsYouTypeReplacePlainTextEmphasis;  // Returns or sets if manual emphasis characters are automatically replaced with character formatting as you type.
@property BOOL autoFormatAsYouTypeReplaceQuotes;  // Returns or sets if straight quotation marks are automatically changed to smart, curly, quotation marks as you type.
@property BOOL autoFormatAsYouTypeReplaceSymbols;  // Returns or sets if two consecutive hyphens -- are replaced with an en dash Ôøê or an em dash Ôøë as you type.
@property BOOL autoFormatDeleteAutoSpaces;  // Returns or sets if Microsoft Word deletes spaces inserted between Japanese and Latin text, when Word formats a document or range automatically.
@property BOOL autoFormatMatchParentheses;  // Returns or sets if Microsoft Word  automatically corrects improperly paired parentheses.
@property BOOL autoFormatPreserveStyles;  // Returns or sets if previously applied styles are preserved when Word formats a document or range automatically.
@property BOOL autoFormatReplaceEastAsianDashes;  // Returns or sets if Microsoft Word automatically corrects long vowel sounds and dashes.
@property BOOL autoFormatReplaceFirstIndents;  // Returns or sets for Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent.
@property BOOL autoFormatReplaceFractions;  // Returns or sets if typed fractions are replaced with fractions from the current character set when Word formats a document or range automatically.
@property BOOL autoFormatReplaceHyperlinks;  // Returns or sets if e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLS, are changed to hyperlinks, when Word formats a document or range automatically.
@property BOOL autoFormatReplaceOrdinals;  // Returns or sets if the ordinal number suffixes st, nd, rd, and th are replaced with the same letters in superscript when Word formats a document or range automatically. For example, 1st is replaced with 1 followed by st formatted as superscript.
@property BOOL autoFormatReplacePlainTextEmphasis;  // Returns or sets if manual emphasis characters are replaced with character formatting when Word formats a document or range automatically.
@property BOOL autoFormatReplaceQuotes;  // Returns or sets if straight quotation marks are automatically changed to smart, curly, quotation marks when Word formats a document or range automatically.
@property BOOL autoFormatReplaceSymbols;  // Returns or set if two consecutive hyphens -- are replaced by an en dash Ôøê or an em dash Ôøë when Word formats a document or range automatically.
@property BOOL autoWordSelection;  // Returns or sets if dragging selects one word at a time instead of one character at a time.
@property BOOL automaticallyBuildUpEquations;  // Returns or sets whether Microsoft Word automatically converts equations to professional format.
@property BOOL ayMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL blueScreen;  // Returns or sets if Word displays text as white characters on a blue background.
@property NSInteger buttonFieldClicks;  // Returns or sets the number of clicks, either one or two, required to run a GOTOBUTTON or MACROBUTTON field.
@property BOOL bvMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL byteMatchFuzzy;  // Returns or sets Microsoft Word ignores the distinction between full-width and half-width characters, Latin or Japanese, during a search.
@property BOOL caseMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between uppercase and lowercase letters during a search.
@property BOOL checkGrammarAsYouType;  // Returns or sets if Word checks grammar and marks errors automatically as you type.
@property BOOL checkGrammarWithSpelling;  // Returns or sets if Word checks grammar while checking spelling.
@property BOOL checkSpellingAsYouType;  // Returns or sets if Word checks spelling and marks errors automatically as you type.
@property WordWdColorIndex commentsColor;  // Returns or sets the color of comments.
@property BOOL confirmConversions;  // Returns or sets if Word displays the convert file dialog box before it opens or inserts a file that isn't a Word document or template. In the convert file dialog box, the user chooses the format to convert the file from.
@property BOOL convertHighAnsiToEastAsian;  // Returns or sets if Microsoft Word converts text that is associated with an East Asian font to the appropriate font when it opens a document.
@property BOOL createBackup;  // Returns or sets if Word creates a backup copy each time a document is saved.
@property BOOL dashMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search.
@property (copy) NSArray<NSNumber *> *defaultBorderColor;  // Returns or sets the default RGB color to use for new border objects.
@property WordWdColorIndex defaultBorderColorIndex;  // Returns or sets the default line color index for borders.
@property WordMsoThemeColorIndex defaultBorderColorThemeIndex;  // Returns or sets the default line color for borders.
@property WordWdLineStyle defaultBorderLineStyle;  // Returns or sets the default border line style
@property WordWdLineWidth defaultBorderLineWidth;  // Returns or sets the default line width of borders.
@property WordWdColorIndex defaultHighlightColorIndex;  // Returns or sets the color index  used to highlight text formatted with the highlight button.
@property WordWdOpenFormat defaultOpenFormat;  // Returns or sets the default file converter used to open documents.
@property WordWdCellColor deletedCellColor;  // Returns or sets the color of cells that are deleted while change tracking is enabled.
@property WordWdColorIndex deletedTextColor;  // Returns or sets the color of text that is deleted while change tracking is enabled.
@property WordWdDeletedTextMark deletedTextMark;  // Returns or sets the format of text that is deleted while change tracking is enabled.
@property BOOL displayGridLines;  // Returns or sets if Microsoft Word displays the document grid. 
@property BOOL displayPasteOptions;  // Returns or sets if Microsoft Word  displays the Paste Options button, which displays directly under newly pasted text.
@property BOOL dzMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL enableSound;  // Returns or sets if Word makes the computer respond with a sound whenever an error occurs.
@property (readonly) BOOL envelopeFeederInstalled;  // Returns true if the current printer has a special feeder for envelopes.
@property double gridDistanceHorizontal;  // Returns or sets the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize autoshapes or East Asian characters in new documents.
@property double gridDistanceVertical;  // Returns or sets the amount of vertical space between the invisible gridlines that Word uses when you draw, move, and resize autoshapes or East Asian characters in new documents.
@property double gridOriginHorizontal;  // Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing autoshapes or East Asian characters to begin in new documents.
@property double gridOriginVertical;  // Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing autoshapes or East Asian characters to begin in new documents.
@property BOOL hfMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL hiraganaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between hiragana and katakana during a search.
@property BOOL ignoreInternetAndFileAddresses;  // Returns or sets if file name extensions,  paths, e-mail addresses, server and share names, also known as UNC paths, and Internet addresses, also known as URLs, are ignored while checking spelling.
@property BOOL ignoreMixedDigits;  // Returns or sets if words that contain numbers are ignored while checking spelling.
@property BOOL ignoreUppercase;  // Returns or sets if words in all uppercase letters are ignored while checking spelling.
@property BOOL inlineConversion;  // Returns or sets if Microsoft Word displays an unconfirmed character string in the Japanese Input Method Editor as an insertion between existing character strings.
@property BOOL insertKeyForPaste;  // Returns or sets if the insert key can be used for pasting the clipboard contents.
@property WordWdCellColor insertedCellColor;  // Returns or sets the color of cells that are inserted while change tracking is enabled.
@property WordWdColorIndex insertedTextColor;  // Returns or sets the color of text that is inserted while change tracking is enabled.
@property WordWdInsertedTextMark insertedTextMark;  // Returns or sets how Microsoft Word formats inserted text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the inserted text color property to control the look of inserted text.
@property BOOL iterationMarkMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between types of repetition marks during a search.
@property BOOL kanjiMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between standard and nonstandard kanji ideography during a search.
@property BOOL keepTrackOfFormatting;  // Returns or sets if Microsoft Word keeps track of formatting.
@property BOOL kiKuMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL liveWordCount;  // Returns or sets if the word count is displayed in the status bar.
@property BOOL mapPaperSize;  // Returns or sets if documents formatted for another country's or region's standard paper size, for example, A4 are automatically adjusted so that they're printed correctly on your country's/region's standard paper size, for example, Letter.
@property WordWdMeasurementUnits measurementUnit;  // Returns or sets the standard measurement unit for Microsoft Word.
@property WordWdCellColor mergedCellColor;  // Returns or sets the color of cells that are merged while change tracking is enabled.
@property WordWdColorIndex moveFromTextColor;  // Returns or sets the color of text that is moved from while change tracking is enabled.
@property WordWdMoveFromTextMark moveFromTextMark;  // Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text.
@property WordWdColorIndex moveToTextColor;  // Returns or sets the color of text that is moved to while change tracking is enabled.
@property WordWdMoveToTextMark moveToTextMark;  // Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text.
@property BOOL oldKanaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between new kana and old kana characters during a search.
@property BOOL overtype;  // Returns or sets if overtype mode is active. In overtype mode, the characters you type replace existing characters one by one. When overtype isn't active, the characters you type move existing text to the right.
@property BOOL pagination;  // Returns or sets if Microsoft Word repaginates documents in the background.
@property BOOL pasteAdjustParagraphSpacing;  // Returns or sets if Microsoft Word automatically adjusts the spacing of paragraphs when cutting and pasting selections.
@property BOOL pasteAdjustTableFormatting;  // Returns or sets if Microsoft Word automatically adjusts the formatting of tables when cutting and pasting selections.
@property BOOL pasteAdjustWordSpacing;  // Returns or sets if Microsoft Word automatically adjusts the spacing of words when cutting and pasting selections.
@property BOOL pasteMergeFromExcel;  // Returns or sets if text formatting will be merged when pasting from Microsoft Excel.
@property BOOL pasteMergeFromPowerPoint;  // Returns or sets if text formatting will be merged when pasting from Microsoft PowerPoint.
@property BOOL pasteMergeLists;  // Returns or sets if the formatting of pasted lists will be merged with surrounding lists.
@property BOOL pasteSmartCutPaste;  // Returns or sets if Microsoft Word intelligently pastes selections into a document.
@property BOOL pasteSmartStyleBehavior;  // Returns or sets if Microsoft Word intelligently merges styles when pasting a selection from a different document.
@property (copy) NSString *pictureEditor;  // Returns or sets the name of the application to use to edit pictures.
@property WordWdWrapTypeMerged pictureWrapTypes;  // Returns or sets the wrapping that is used to insert or paste pictures.
@property BOOL plainEquationsUsePlainText;  // Returns or sets if equations are represented in plain text; false uses MathML
@property BOOL printComments;  // Returns or sets if Microsoft Word prints comments, starting on a new page at the end of the document.
@property BOOL printDrawingObjects;  // Returns or sets if Microsoft Word prints drawing objects.
@property BOOL printFieldCodes;  // Returns or sets if Microsoft Word prints field codes instead of field results.
@property BOOL printHiddenText;  // Returns or sets if hidden text is printed.
@property BOOL printProperties;  // Returns or sets if Microsoft Word prints document summary information on a separate page at the end of the document. 
@property BOOL printReverse;  // Returns or sets if Microsoft Word prints pages in reverse order.
@property BOOL prolongedSoundMarkMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between short and long vowel sounds during a search.
@property BOOL punctuationMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between types of punctuation marks during a search.
@property BOOL replaceSelection;  // Returns or sets if the result of typing or pasting replaces the selection. If false the result of typing or pasting is added before the selection, leaving the selection intact.
@property WordWdColorIndex revisedLinesColor;  // Returns or sets the color of changed lines in a document with tracked changes.
@property WordWdRevisedLinesMark revisedLinesMark;  // Returns or sets the placement of changed lines in a document with tracked changes.
@property WordWdColorIndex revisedPropertiesColor;  // Returns or sets the color index used to mark formatting changes while change tracking is enabled. 
@property WordWdRevisedPropertiesMark revisedPropertiesMark;  // Returns or sets the mark used to show formatting changes while change tracking is enabled.
@property NSInteger saveInterval;  // Returns or sets the time interval in minutes for saving autorecover information.
@property BOOL saveNormalPrompt;  // Returns or sets if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it quits. if this is set to false Word automatically saves changes to the Normal template before it quits.
@property BOOL savePropertiesPrompt;  // Returns or sets if Microsoft Word prompts for document property information when saving a new document.
@property BOOL sendMailAttach;  // True if the send to command on the file menu inserts the active document as an attachment to a mail message. False if the send to command inserts the contents of the active document as text in a mail message.
@property BOOL showReadabilityStatistics;  // Returns or sets if Microsoft Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar.
@property BOOL smallKanaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between diphthongs and double consonants during a search.
@property BOOL smartCutPaste;  // Returns or sets if Microsoft Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs.
@property BOOL smartParagraphSelection;  // Returns or sets if Microsoft Word includes the paragraph mark in a selection when selecting most or all of a paragraph.
@property BOOL snapToGrid;  // Returns or sets if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in new documents.
@property BOOL snapToShapes;  // Returns or sets if Word automatically aligns autoshapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other autoshapes or East Asian characters in new documents.
@property BOOL spaceMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between space markers used during a search.
@property WordWdCellColor splitCellColor;  // Returns or sets the color of cells that are split while change tracking is enabled.
@property BOOL suggestFromMainDictionaryOnly;  // Returns or sets if Microsoft Word draws spelling suggestions from the main dictionary only. If false it draws spelling suggestions from the main dictionary and any custom dictionaries that have been added.
@property BOOL suggestSpellingCorrections;  // Returns or sets if Microsoft Word always suggests alternative spellings for each misspelled word when checking spelling.
@property BOOL tabIndentKey;  // Returns or sets if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered and centered paragraphs to left-aligned.
@property BOOL tcMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL updateFieldsAtPrint;  // Returns or sets if Microsoft Word updates fields automatically before printing a document. 
@property BOOL updateLinksAtOpen;  // Returns or sets if Microsoft Word automatically updates all embedded OLE links in a document when it's opened.
@property BOOL updateLinksAtPrint;  // Returns or sets if Microsoft Word updates fields automatically before printing a document.
@property BOOL useCharacterUnit;  // Returns or sets if Microsoft Word uses characters as the default measurement unit for the current document.
@property BOOL useGermanSpellingReform;  // Returns or sets if Microsoft Word uses the German post-reform spelling rules when checking spelling.
@property BOOL warnBeforeSavingPrintingSendingMarkup;  // Returns or sets if Microsoft Word displays a warning when saving, printing, or sending as e-mail a document containing comments or tracked changes.
@property BOOL zjMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.


@end

// Represents a single add-in, either installed or not installed.
@interface WordAddIn : SBObject <WordGenericMethods>

@property (readonly) BOOL autoload;  // Returns true if the specified add-in is automatically loaded when Word is started.
@property (readonly) BOOL compiled;  // Returns true if the specified add-in is a Word add-in library. False if the add-in is a template.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property BOOL installed;  // Returns or sets if the specified add-in is installed.
@property (copy, readonly) NSString *name;  // Returns the name of the add in.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified add-in in HFS path style.


@end

// An AppleScript object representing the Microsoft Word application.
@interface WordApplication (MicrosoftWordSuite)

- (SBElementArray<WordDocument *> *) documents;
- (SBElementArray<WordWindow *> *) windows;
- (SBElementArray<WordRecentFile *> *) recentFiles;
- (SBElementArray<WordFileConverter *> *) fileConverters;
- (SBElementArray<WordCaptionLabel *> *) captionLabels;
- (SBElementArray<WordAddIn *> *) addIns;
- (SBElementArray<WordCommandBar *> *) commandBars;
- (SBElementArray<WordTemplate *> *) templates;
- (SBElementArray<WordKeyBinding *> *) keyBindings;
- (SBElementArray<WordDictionary *> *) dictionaries;

@property (copy, readonly) WordDocument *activeDocument;  // Returns the currently active document object.
@property (copy) NSString *activePrinter;  // Returns or sets the name of the active printer
@property (copy, readonly) WordWindow *activeWindow;  // Returns the currently active window object.
@property BOOL animations;  // Returns or sets if animations are enabled while a script is running.
@property (copy, readonly) NSString *applicationVersion;  // Returns the Microsoft Word version number as a string.
@property (copy, readonly) WordAutocorrect *autocorrectObject;  // Returns the autocorrect object
@property (copy) id automationSecurity;
@property (readonly) NSInteger backgroundPrintingStatus;  // Returns the number of print jobs in the background printing queue.
@property (copy) NSString *browseExtraFileTypes;  // Returns or sets to open hyperlinked HTML files in the Internet browser or Microsoft Word.  Set this property to text/html to allow hyperlinked HTML files to be opened in Microsoft Word.
@property (copy, readonly) WordBrowser *browserObject;  // Returns the browser object.  The browser object is a tool used to move the insertion point to objects in a document.
@property (copy, readonly) NSString *build;  // Returns the version and build number of the Microsoft Word application.
@property (readonly) BOOL capsLock;  // Returns if caps lock is turned on.
@property (copy, readonly) NSString *caption;  // Returns the name of the application.
@property (copy) SBObject *customizationContext;  // Returns or set a template or document object that represents the template or document in which changes to menus, toolbars, and key bindings are stored.
@property (copy) NSString *defaultSaveFormat;  // Returns or sets the default format. Common settings include: document = WordDocument, document template = Template, Word 97-2004 document = Doc97, Word XML document = XML, web page = Html, Text only = Text, RTF = Rtf, unicode text = Unicode.
@property (copy) NSString *defaultStartupPath;  // Returns or sets the complete path of the startup folder, excluding the final separator.
@property (copy) NSString *defaultTableSeparator;  // Returns or sets the single character used to separate text into cells when text is converted to a table.
@property (copy, readonly) WordDefaultWebOptions *defaultWebOptionsObject;  // Returns the default web options object.
@property WordWdAlertLevel displayAlerts;  // Returns or sets the way certain alerts or messages are handled while an AppleScript is running.
@property BOOL displayAutoCompleteTips;  // Returns or set if Microsoft Word displays tips that suggest text for completing words, dates, or phrases as you type.
@property BOOL displayRecentFiles;  // Returns or sets if the names of recently used files are displayed on the file menu.
@property BOOL displayScreenTips;  // Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted.
@property BOOL displayScrollBars;  // Returns or sets if Word displays a scroll bar in at least one document window.  Setting this property to true will display horizontal and vertical scrollbars in all windows.  Setting this property to false turns off all scrollbars in all windows.
@property BOOL displayStatusBar;  // Returns or sets if the status bar is displayed.
@property BOOL doPrintPreview;  // Returns or set if print preview is the current view.
@property (copy, readonly) NSArray<NSString *> *fontNames;  // Returns the list of names of all of the available fonts
@property NSInteger keyboardScript;  // Returns or sets the current keyboard script
@property (copy, readonly) NSArray<NSString *> *landscapeFontNames;  // Returns the list of names of all of the available landscape fonts
@property (readonly) NSInteger localizedLanguage;  // Returns the Windows language ID for the locale that Microsoft Word is using.
@property (copy, readonly) WordMailingLabel *mailingLabelObject;  // Returns the mailing label object.
@property (copy, readonly) WordMathAutocorrect *mathAc;  // Returns a math autocorrect object that represents the auto correct entries for equations.
@property (copy, readonly) NSString *name;  // Returns the name of this application.
@property (copy, readonly) WordTemplate *normalTemplate;  // Returns the normal template object
@property (readonly) BOOL numLock;  // Returns the state of the num lock key.  True if the keys on the numeric keypad insert numbers, false if the keys move the insertion point.
@property (copy, readonly) NSString *path;  // Returns the path to the application in HFS path style
@property (copy, readonly) NSString *pathSeparator;  // Returns the character used to separate folder names.
@property (copy, readonly) NSArray<NSString *> *portraitFontNames;  // Returns the list of names of all of the available portrait fonts
@property (copy, readonly) WordSelectionObject *selection;  // Returns the selection object.
@property (copy, readonly) WordWordOptions *settings;  // Returns the word options object.
@property BOOL showVisualBasicEditor;  // Return or set if the visual basic editor is visible.
@property (readonly) BOOL specialMode;  // Returns if Microsoft Word is in a special mode for example, copy text mode or move text mode.
@property BOOL startupDialog;  // Returns of sets if the project gallery dialog will be shown when starting Microsoft Word.
@property (copy) NSString *statusBar;  // Displays the specified text in the status bar. This is a write only property.
@property (copy, readonly) WordSystemObject *system_object;  // Returns the system object
@property (readonly) NSInteger usableHeight;  // Returns the maximum height in points to which you can set the width of a Microsoft Window document window.
@property (readonly) NSInteger usableWidth;  // Returns the maximum width in pixels to which you can set the width of a Microsoft Window document window.
@property (copy) NSString *userAddress;  // Returns of set the users mailing address.
@property (readonly) BOOL userControl;  // Returns true if the application was launched by a user.  False if the program was opened programmatically.
@property (copy) NSString *userInitials;  // Returns of sets the initials, which Microsoft Word uses to construct comment marks.
@property (copy) NSString *userName;  // Returns or sets the users name, which is used on envelopes and for the author document property.

@end

// Represents a single AutoText entry.
@interface WordAutoTextEntry : SBObject <WordGenericMethods>

@property (copy) NSString *autoTextValue;  // Returns or sets the value of this auto text entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of this auto text entry.
@property (copy, readonly) NSString *styleName;  // Returns the name of the style applied to the specified auto text entry.

- (WordTextRange *) insertAutoTextEntryWhere:(WordTextRange *)where richText:(BOOL)richText;  // Inserts the auto text entry in place of the specified range. Returns a text range object that represents the auto text entry.

@end

// Represents a single bookmark.
@interface WordBookmark : SBObject <WordGenericMethods>

@property (readonly) BOOL column;  // True if the specified bookmark is a table column.
@property (readonly) BOOL empty;  // True if the specified bookmark is empty. An empty bookmark marks a location of a collapsed selection, it doesn't mark any text.
@property NSInteger endOfBookmark;  // Returns or sets the ending character position of the bookmark.
@property (copy, readonly) NSString *name;  // The name of the bookmark object.
@property NSInteger startOfBookmark;  // Returns or sets the starting character position of the bookmark.
@property (readonly) WordWdStoryType storyType;  // Returns the story type for the bookmark.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's referred to by the bookmark.

- (WordBookmark *) copyBookmarkName:(NSString *)name NS_RETURNS_NOT_RETAINED;  // Sets the bookmark specified by the name argument to the location marked by another bookmark, and returns a bookmark object.

@end

// Represents options associated with border object.
@interface WordBorderOptions : SBObject <WordGenericMethods>

@property BOOL alwaysInFront;  // Returns or sets if page borders are displayed in front of the document text.
@property WordWdBorderDistanceFrom distanceFrom;  // Returns or sets a value that indicates whether the specified page border is measured from the edge of the page or from the text it surrounds.
@property NSInteger distanceFromBottom;  // Returns or sets the space in points between the text and the bottom border.
@property NSInteger distanceFromLeft;  // Returns or sets the space in points between the text and the left border.
@property NSInteger distanceFromRight;  // Returns or sets the space in points between the right edge of the text and the right border. 
@property NSInteger distanceFromTop;  // Returns or sets the space in points between the text and the top border.
@property BOOL enableBorders;  // Returns or sets border formatting for the specified object.
@property BOOL enableFirstPageInSection;  // Returns or sets if page borders are enabled for the first page in the section.
@property BOOL enableOtherPagesInSection;  // Returns or sets if page borders are enabled for all pages in the section except for the first page. 
@property (readonly) BOOL hasHorizontal;  // Returns true if a horizontal border can be applied to the object. 
@property (readonly) BOOL hasVertical;  // Returns true if a vertical border can be applied to the object. 
@property (copy) NSArray<NSNumber *> *insideColor;  // Returns or sets the RGB color of the inside borders
@property WordWdColorIndex insideColorIndex;  // Returns or sets the color index of the inside borders. 
@property WordMsoThemeColorIndex insideColorThemeIndex;  // Returns or sets the color of the inside borders.
@property WordWdLineStyle insideLineStyle;  // Returns or sets the inside border for the specified object.
@property WordWdLineWidth insideLineWidth;  // Returns or sets the line width of the inside border of an object.
@property BOOL joinBorders;  // Returns or sets if vertical borders at the edges of paragraphs and tables are removed so that the horizontal borders can connect to the page border.
@property (copy) NSArray<NSNumber *> *outsideColor;  // Returns or sets the RGB color of the outside borders
@property WordWdColorIndex outsideColorIndex;  // Returns or sets the color index of the outside borders. 
@property WordMsoThemeColorIndex outsideColorThemeIndex;  // Returns or sets the color of the outside borders.
@property WordWdLineStyle outsideLineStyle;  // Returns or sets the outside border for the specified object.
@property WordWdLineWidth outsideLineWidth;  // Returns or sets the line width of the outside border of an object.
@property BOOL shadow;  // Returns or sets if the specified border is formatted as shadowed.
@property BOOL surroundFooter;  // Returns or sets if a page border encompasses the document footer.
@property BOOL surroundHeader;  // Returns or sets if a page border encompasses the document header.

- (void) applyPageBordersToAllSections;  // Applies the specified page-border formatting to all sections in a document.

@end

// Represents a border of an object.
@interface WordBorder : SBObject <WordGenericMethods>

@property WordWdPageBorderArt artStyle;  // Returns or sets the graphical page-border design for a document.
@property NSInteger artWidth;  // Returns or sets the width in points of the specified graphical page border.
@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the RGB color for the specified border object. 
@property WordWdColorIndex colorIndex;  // Returns or sets the color index for the specified border.
@property WordMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified border or font object.
@property (readonly) BOOL inside;  // Returns true if an inside border can be applied to the specified object.
@property WordWdLineStyle lineStyle;  // Returns or sets the border line style for the specified object.
@property WordWdLineWidth lineWidth;  // Returns or sets the line width of an object's border.
@property BOOL visible;  // Returns or sets if the border object is visible


@end

// Represents the browser tool used to move the insertion point to objects in a document. This tool is comprised of the three buttons at the bottom of the vertical scroll bar.
@interface WordBrowser : SBObject <WordGenericMethods>

@property WordWdBrowseTarget browserTarget;  // Returns or sets the document item that the previous and next methods locate.

- (void) nextForBrowser;  // Moves the selection to the next item indicated by the browser target. Use the browser target property to change the browser target.
- (void) previousForBrowser;  // Moves the selection to the previous item indicated by the browser target. Use the browser target property to change the browser target.

@end

@interface WordBuildingBlockCategory : SBObject <WordGenericMethods>

- (SBElementArray<WordBuildingBlock *> *) buildingBlocks;

@property (copy, readonly) WordBuildingBlockType *buildingBlockType;  // Returns a building block type object that represents the type of building block for a building block category.
@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *name;  // Returns the name of the specified object.

- (WordBuildingBlock *) addBuildingBlockToCategoryName:(NSString *)name fromLocation:(WordTextRange *)fromLocation objectDescription:(NSString *)objectDescription insertOptions:(WordWdDocPartInsertOptions)insertOptions;  // Creates a new building block.

@end

// Represents a type of building block.
@interface WordBuildingBlockType : SBObject <WordGenericMethods>

- (SBElementArray<WordBuildingBlockCategory *> *) buildingBlockCategories;

@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *name;  // Returns a String that represents the localized name of a building block type.


@end

@interface WordBuildingBlock : SBObject <WordGenericMethods>

@property (copy, readonly) WordBuildingBlockType *buildingBlockType;  // Returns a building block type object that represents the type for a building block.
@property (copy, readonly) WordBuildingBlockCategory *category;  // Returns a Category object that represents the category for a building block.
@property (copy) NSString *objectDescription;  // Gets or sets a string that represents the description for a building block.
@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *identifier;  // Returns a string that represents the internal identification number for a building block.
@property NSInteger insertOptions;  // Gets or sets an integer that represents how to insert the contents of a building block into a document.
@property (copy) NSString *name;  // Gets or sets a string that represents the name of a building block.
@property (copy) NSString *value;  // Gets or sets a string that represents the contents of a building block.

- (WordTextRange *) insertBuildingBlockInLocation:(WordTextRange *)inLocation asRichText:(BOOL)asRichText;  // Inserts the value of a building block into a document and returns a text range object that represents the contents of the building block within the document.

@end

// Represents a single caption label.
@interface WordCaptionLabel : SBObject <WordGenericMethods>

@property (readonly) BOOL builtIn;  // Returns true if this is a built-in caption label.
@property (readonly) WordWdCaptionLabelID captionLabelId;  // Returns a constant that represents the type for the specified caption label if the built in property of the caption label object is true.
@property WordWdCaptionPosition captionLabelPosition;  // Returns or sets the position of caption label text.
@property NSInteger chapterStyleLevel;  // Returns or sets the heading style that marks a new chapter when chapter numbers are included with the specified caption label. The number 1 corresponds to Heading 1, 2 corresponds to Heading 2, and so on.
@property BOOL includeChapterNumber;  // Returns or sets if a chapter number is included with page numbers or a caption label.
@property (copy, readonly) NSString *name;  // Returns the name for the caption label
@property WordWdCaptionNumberStyle numberStyle;  // Returns or sets the number style for the caption label object.
@property WordWdSeparatorType separator;  // Returns or sets the character between the chapter number and the sequence number.


@end

// Represents a single check box form field.
@interface WordCheckBox : SBObject <WordGenericMethods>

@property BOOL autoSize;  // True sizes the check box according to the font size of the surrounding text. False sizes the check box according to the size property.
@property BOOL checkBoxDefault;  // Returns or sets the default check box value. True if the default value is checked. 
@property BOOL checkBoxValue;  // Returns or sets if the check box is selected.  True if the check box is selected.
@property double checkboxSize;  // Returns or sets the size of the specified check box in points.
@property (readonly) BOOL valid;  // Returns if the check box object is valid.


@end

// Represents a co-authoring lock
@interface WordCoauthLock : SBObject <WordGenericMethods>

@property BOOL isHeader_footer_lock;  // returns if this lock is a header/footer lock.
@property (copy, readonly) WordCoauthor *lockOwner;  // Get the owner of the co-authoring lock.
@property WordWdLockType lockType;  // Get the type of the co-authoring lock.
@property (copy, readonly) WordTextRange *textObject;  // Gets a text range object that represents the portion of a document that contains the co-authoring lock.

- (void) unlock;  // Remove the co-authoring lock.

@end

// Represents a co-authoring update
@interface WordCoauthUpdate : SBObject <WordGenericMethods>

@property (copy, readonly) WordTextRange *textObject;  // Gets a text range object that represents the portion of a document that contains the co-authoring update.


@end

// Represents a coauthor
@interface WordCoauthor : SBObject <WordGenericMethods>

- (SBElementArray<WordCoauthLock *> *) coauthLocks;

@property (copy, readonly) NSString *coauthorEmailAddress;  // The email address of coauthor.
@property (copy, readonly) NSString *coauthorId;  // The ID of coauthor.
@property (copy, readonly) NSString *coauthorName;  // The name of coauthor.
@property BOOL isMe;  // returns if this coauthor is me.


@end

// Represents a coauthoring
@interface WordCoauthoring : SBObject <WordGenericMethods>

- (SBElementArray<WordCoauthor *> *) coauthors;
- (SBElementArray<WordCoauthLock *> *) coauthLocks;
- (SBElementArray<WordCoauthUpdate *> *) coauthUpdates;
- (SBElementArray<WordConflict *> *) conflicts;

@property BOOL canMerge;  // returns if coauthoring can merge.
@property BOOL canShare;  // returns if coauthoring can share.
@property (copy) WordCoauthor *myself;  // returns me as author.
@property BOOL pendingUpdates;  // returns if any updates are pending.


@end

// Represents a conflict
@interface WordConflict : SBObject <WordGenericMethods>

@property (readonly) WordWdRevisionType conflictType;  // Returns the revision type.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the conflict


@end

// Represents a custom mailing label.
@interface WordCustomLabel : SBObject <WordGenericMethods>

@property (readonly) BOOL dotMatrix;  // True if the printer type for the specified custom label is dot matrix. False if the printer type is either laser or ink jet.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property double height;  // Returns or sets the height of the object.
@property double horizontalPitch;  // Returns or sets the horizontal distance in points between the left edge of one custom mailing label and the left edge of the next mailing label.
@property (copy) NSString *name;  // Returns or set the name of the custom mailing label.
@property NSInteger numberAcross;  // Returns or sets the number of custom mailing labels across a page.
@property NSInteger numberDown;  // Returns or sets the number of custom mailing labels down the length of a page.
@property WordWdCustomLabelPageSize pageSize;  // Returns or sets the page size for the specified custom mailing label.
@property double sideMargin;  // Returns or sets the side margin widths in points for the specified custom mailing label.
@property double topMargin;  // Returns or sets the distance in points between the top edge of the page and the top boundary of the body text.
@property (readonly) BOOL valid;  // True if the various properties for example, height, width, and number down for the specified custom label work together to produce a valid mailing label.
@property double verticalPitch;  // Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label.
@property double width;  // Returns or sets the width of the object.


@end

// Represents a single data merge field in a data source.
@interface WordDataMergeDataField : SBObject <WordGenericMethods>

@property (copy, readonly) NSString *dataMergeDataFieldValue;  // Returns the contents of the mail merge data field or mapped data field for the current record. Use the active record property to set the active record in a data merge data source.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the data merge data field object.


@end

// Represents the data merge data source in a data merge operation.
@interface WordDataMergeDataSource : SBObject <WordGenericMethods>

- (SBElementArray<WordDataMergeFieldName *> *) dataMergeFieldNames;
- (SBElementArray<WordDataMergeDataField *> *) dataMergeDataFields;

@property WordWdMailMergeActiveRecord activeRecord;  // Returns back the index of the current active record or an enumeration specifying the record
@property (copy, readonly) NSString *connectString;  // Returns the connection string for the specified data merge data source.
@property NSInteger firstRecord;  // Returns or sets the number of the first data record to be merged in a data merge operation. 
@property (copy, readonly) NSString *headerSourceName;  // Returns the path and file name of the header source attached to the specified mail merge main document.
@property (readonly) WordWdMailMergeDataSource headerSourceType;  // Returns a value that indicates the way the header source is being supplied for the mail merge operation.
@property NSInteger lastRecord;  // Returns or sets the number of the last data record to be merged in a data merge operation. 
@property (readonly) WordWdMailMergeDataSource mailMergeDataSourceType;  // Returns the type of data merge data source.
@property (copy, readonly) NSString *name;  // Returns the name of the data merge data source.
@property (copy) NSString *queryString;  // Returns or sets the query string, SQL statement, used to retrieve a subset of the data in a data merge data source.

- (BOOL) findRecordFindText:(NSString *)findText fieldName:(NSString *)fieldName;  // Searches the contents of the specified data merge data source for text in a particular field. Returns true if the search text is found.

@end

// Represents a data merge field name in a data source.
@interface WordDataMergeFieldName : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the data merge field name object


@end

// Represents a single data merge field in a document.
@interface WordDataMergeField : SBObject <WordGenericMethods>

@property (copy) WordTextRange *dataMergeFieldRange;  // Returns or sets a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters.
@property (readonly) WordWdFieldType formFieldType;  // The type of this data merge field
@property BOOL locked;  // Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results.
@property (copy, readonly) WordDataMergeField *nextDataMergeField;  // Returns the next data merge field
@property (copy, readonly) WordDataMergeField *previousMakeMergeField;  // Returns the previous data merge field


@end

// Represents the data merge functionality in Word.
@interface WordDataMerge : SBObject <WordGenericMethods>

- (SBElementArray<WordDataMergeField *> *) dataMergeFields;

@property (copy, readonly) WordDataMergeDataSource *dataSource;  // Returns or sets the destination of the data merge results.
@property WordWdMailMergeDestination destination;  // Returns or sets the destination of the data merge results.
@property (copy) NSString *mailAddressFieldName;  // Returns or sets the name of the field that contains e-mail addresses that are used when the data merge destination is electronic mail.
@property BOOL mailAsAttachment;  // Returns or sets if the merge documents are sent as attachments when the data merge destination is an e-mail message or a fax.
@property (copy) NSString *mailSubject;  // Returns or sets the subject line used when the data merge destination is electronic mail.
@property WordWdMailMergeMainDocType mainDocumentType;  // Returns or sets the data merge main document type.
@property (readonly) WordWdMailMergeState state;  // Returns the current state of a data merge operation.
@property BOOL suppressBlankLines;  // Returns or sets if blank lines are suppressed when data merge fields in a data merge main document are empty.
@property BOOL viewDataMergeFieldCodes;  // Returns or sets if merge field names are displayed in a data merge main document. False if information from the current data record is displayed. 

- (void) check;  // Simulates the data merge operation, pausing to report each error as it occurs.
- (void) createDataSourceName:(NSString *)name passwordDocument:(NSString *)passwordDocument writePassword:(NSString *)writePassword headerRecord:(NSString *)headerRecord MSQuery:(BOOL)MSQuery SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1 connection:(NSString *)connection linkToSource:(BOOL)linkToSource;  // Creates a Word document that uses a table to store data for a data merge. The new data source is attached to the specified document, which becomes a main document if it's not one already.
- (void) createHeaderSourceName:(NSString *)name passwordDocument:(NSString *)passwordDocument writePassword:(NSString *)writePassword headerRecord:(NSString *)headerRecord;  // Creates a Word document that stores a header record that's used in place of the data source header record in a mail merge. This method attaches the new header source to the specified document, which becomes a main document if it's not one already.
- (void) editDataSource;  // Opens or switches to the mail merge data source.
- (void) editHeaderSource;  // Opens the header source attached to a mail merge main document, or activates the header source if it's already open.
- (void) editMainDocument;  // Activates the mail merge main document associated with the specified header source or data source document.
- (void) executeDataMergePause:(BOOL)pause;  // Performs the specified data merge operation. Performs the specified data merge operation.
- (WordDataMergeField *) makeNewDataMergeAskFieldTextRange:(WordTextRange *)textRange name:(NSString *)name prompt:(NSString *)prompt defaultAskText:(NSString *)defaultAskText askOnce:(BOOL)askOnce;  // Create a new data merge ask field
- (WordDataMergeField *) makeNewDataMergeFillInFieldTextRange:(WordTextRange *)textRange prompt:(NSString *)prompt defaultFillInText:(NSString *)defaultFillInText askOnce:(BOOL)askOnce;  // Create a new data merge fill in field
- (WordDataMergeField *) makeNewDataMergeIfFieldTextRange:(WordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(WordWdMailMergeComparison)comparison compareTo:(NSString *)compareTo trueText:(NSString *)trueText falseText:(NSString *)falseText;  // Create a new data merge if field
- (WordDataMergeField *) makeNewDataMergeNextFieldTextRange:(WordTextRange *)textRange;  // Create a new data merge next field
- (WordDataMergeField *) makeNewDataMergeNextIfFieldTextRange:(WordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(WordWdMailMergeComparison)comparison compareTo:(NSString *)compareTo;  // Create a new data merge next if field
- (WordDataMergeField *) makeNewDataMergeRecFieldTextRange:(WordTextRange *)textRange;  // Create a new data merge rec field
- (WordDataMergeField *) makeNewDataMergeSequenceFieldTextRange:(WordTextRange *)textRange;  // Create a new data merge sequence field
- (WordDataMergeField *) makeNewDataMergeSetFieldTextRange:(WordTextRange *)textRange name:(NSString *)name valueText:(NSString *)valueText;  // Create a new data merge set field
- (WordDataMergeField *) makeNewDataMergeSkipIfFieldTextRange:(WordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(WordWdMailMergeComparison)comparison compareTo:(NSString *)compareTo;  // Create a new data merge skip if field
- (void) openDataSourceName:(NSString *)name format:(WordWdOpenFormat)format confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly linkToSource:(BOOL)linkToSource addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate connection:(NSString *)connection SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1;  // Attaches a data source to the specified document, which becomes a main document if it's not one already.
- (void) openHeaderSourceName:(NSString *)name format:(WordWdOpenFormat)format confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate;  // Attaches a mail merge header source to the specified document.
- (void) useAddressBookBookType:(NSString *)bookType;  // Selects the address book that's used as the data source for a mail merge operation.

@end

// Contains global application-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page. You can return or set attributes either at the application global level or at the document level.
@interface WordDefaultWebOptions : SBObject <WordGenericMethods>

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page.
@property BOOL alwaysSaveInDefaultEncoding;  // Returns or saves if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.  The default value is False.
@property BOOL checkIfOfficeIsHtmleditor;  // Returns or sets if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word.
@property BOOL checkIfWordIsDefaultHtmleditor;  // Returns or sets if Microsoft Word checks to see whether it is the default HTML editor when you start Word. The default value is true.
@property WordMsoEncoding encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property WordMsoScreenSize screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL updateLinksOnSave;  // Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved.
@property BOOL useLongFileNames;  // Returns or sets if long file names are used when you save the document as a Web page.


@end

// Represents a built-in dialog box.
@interface WordDialog : SBObject <WordGenericMethods>

@property WordWdWordDialogTab defaultDialogTab;  // Returns or sets the active tab when the specified dialog box is displayed.
@property (readonly) WordWdWordDialog dialogType;  // The built-in dialog this object represents.

- (NSInteger) displayWordDialogTimeOut:(NSInteger)timeOut;  // Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a Long that indicates which button was clicked to close the dialog box.
- (void) executeDialog;  // Applies the current settings of a Microsoft Word dialog box.
- (NSInteger) showTimeOut:(NSInteger)timeOut;  // Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a number which indicates the button used to dismiss the dialog box.

@end

// Represents a single version of a document.
@interface WordDocumentVersion : SBObject <WordGenericMethods>

@property (copy, readonly) NSString *comment;  // Returns the comment associated with the specified version of a document.
@property (copy, readonly) NSDate *dateValue;  // The date and time that the document version was saved.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *savedBy;  // Returns the name of the user who saved the specified version of the document.

- (void) openVersion;  // Opens the specified version of the document.

@end

// Represents a Microsoft Word document.
@interface WordDocument (MicrosoftWordSuite)

- (SBElementArray<WordDocumentProperty *> *) documentProperties;
- (SBElementArray<WordCustomDocumentProperty *> *) customDocumentProperties;
- (SBElementArray<WordBookmark *> *) bookmarks;
- (SBElementArray<WordTable *> *) tables;
- (SBElementArray<WordFootnote *> *) footnotes;
- (SBElementArray<WordEndnote *> *) endnotes;
- (SBElementArray<WordWordComment *> *) WordComments;
- (SBElementArray<WordSection *> *) sections;
- (SBElementArray<WordParagraph *> *) paragraphs;
- (SBElementArray<WordWord *> *) words;
- (SBElementArray<WordSentence *> *) sentences;
- (SBElementArray<WordCharacter *> *) characters;
- (SBElementArray<WordField *> *) fields;
- (SBElementArray<WordFormField *> *) formFields;
- (SBElementArray<WordWordStyle *> *) WordStyles;
- (SBElementArray<WordFrame *> *) frames;
- (SBElementArray<WordTableOfFigures *> *) tablesOfFigures;
- (SBElementArray<WordVariable *> *) variables;
- (SBElementArray<WordRevision *> *) revisions;
- (SBElementArray<WordTableOfContents *> *) tablesOfContents;
- (SBElementArray<WordTableOfAuthorities *> *) tablesOfAuthorities;
- (SBElementArray<WordWindow *> *) windows;
- (SBElementArray<WordIndex *> *) indexes;
- (SBElementArray<WordSubdocument *> *) subdocuments;
- (SBElementArray<WordStoryRange *> *) storyRanges;
- (SBElementArray<WordHyperlinkObject *> *) hyperlinkObjects;
- (SBElementArray<WordShape *> *) shapes;
- (SBElementArray<WordListTemplate *> *) listTemplates;
- (SBElementArray<WordWordList *> *) WordLists;
- (SBElementArray<WordInlineShape *> *) inlineShapes;
- (SBElementArray<WordDocumentVersion *> *) documentVersions;
- (SBElementArray<WordReadabilityStatistic *> *) readabilityStatistics;
- (SBElementArray<WordGrammaticalError *> *) grammaticalErrors;
- (SBElementArray<WordSpellingError *> *) spellingErrors;
- (SBElementArray<WordMathObject *> *) mathObjects;

@property (copy, readonly) WordWindow *activeWindow;  // Returns a window object that represents the active window, the window with the focus. If there are no windows open, an error occurs.
@property (copy) WordTemplate *attachedTemplate;  // Returns of sets a template object that represents the template attached to the specified document. To set this property, specify either the name of the template or a template object.
@property BOOL autoHyphenation;  // Returns or sets if automatic hyphenation is turned on for the specified document.
@property (copy) WordShape *backgroundShape;  // Returns or sets a shape object that represents the background image for the specified document.
@property WordWdBuiltinStyle clickAndTypeParagraphStyle;  // Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document.
@property (copy, readonly) WordCoauthoring *coauthoring;
@property WordWdCompatibilityVersion compatibleVersion;  // Returns or sets the compatibility options to a given application version.
@property NSInteger consecutiveHyphensLimit;  // Returns or sets the maximum number of consecutive lines that can end with hyphens. If this property is set to zero, any number of consecutive lines can end with hyphens.
@property (copy, readonly) WordDataMerge *dataMerge;  // Returns the data merge object.
@property double defaultTabStop;  // Returns or sets the interval in points between the default tab stops in the specified document.
@property (readonly) WordWdBuiltinStyle defaultTableStyle;  // Returns the default table style for this document.
@property (copy, readonly) WordOfficeTheme *documentTheme;  // Returns an office theme object that represents the Microsoft Office theme applied to a document.
@property (readonly) WordWdDocumentType document_type;  // Returns the document type.
@property WordWdFarEastLineBreakLevel eastAsianLineBreak;  // Returns or sets the line break control level for the specified document. This property is ignored if the Ôøíeast asian line break controlÔøì property is set to false. Note that Ôøíeast asian line break controlÔøì is a paragraph format property.
@property BOOL embedTrueTypeFonts;  // Returns or set if Microsoft Word embeds TrueType fonts in a document when it's saved. This allow others to view the document with the same fonts that were used to create it.
@property (copy, readonly) WordEndnoteOptions *endnoteOptions;  // Returns the endnote options object.
@property (copy, readonly) WordEnvelope *envelopeObject;  // Returns the envelop object.
@property (copy, readonly) WordFieldOptions *fieldOptions;
@property (copy, readonly) WordFootnoteOptions *footnoteOptions;  // Returns the footnote options object.
@property (copy, readonly) NSString *fullName;  // Returns the full name of the document in HFS path style.
@property BOOL grammarChecked;  // True if a grammar check has been run on the document. False if some of the document hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false.
@property double gridDistanceHorizontal;  // Returns or sets the amount of horizontal space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize shape object or east asian characters in the document.
@property double gridDistanceVertical;  // Returns or sets the amount of vertical space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize shape object or east asian characters in the document.
@property BOOL gridOriginFromMargin;  // Returns or sets if Microsoft Word starts the character grid from the upper-left corner of the page.
@property double gridOriginHorizontal;  // Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing shape object or east asian characters to begin in the document.
@property double gridOriginVertical;  // Returns or sets the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing shape object or east asian characters to begin in the document.
@property NSInteger gridSpaceBetweenHorizontalLines;  // Returns or sets the interval at which Microsoft Word displays horizontal character gridlines in print layout view.
@property NSInteger gridSpaceBetweenVerticalLines;  // Returns or sets the interval at which Microsoft Word displays vertical character gridlines in print layout view.
@property (readonly) BOOL hasPassword;  // True if a password is required to open the specified document.
@property BOOL hyphenateCaps;  // Returns or sets if words in all capital letters can be hyphenated.
@property NSInteger hyphenationZone;  // Returns or sets the width of the hyphenation zone, in points. The hyphenation zone is the maximum amount of space that Microsoft Word leaves between the end of the last word in a line and the right margin.
@property (readonly) BOOL integralsUsesSubscripts;  // Gets or sets a value that specifies the default location of limits for integrals.
@property (readonly) BOOL isMasterDocument;  // True if the specified document is a master document. A master document includes one or more subdocuments.
@property (readonly) BOOL isSubdocument;  // True if the specified document is opened in a separate document window as a subdocument of a master document.
@property WordWdJustificationMode justificationMode;  // Returns or sets the character spacing adjustment for the specified document.
@property (copy) WordLetterContent *letterContent;  // Return or sets the letter content object associated with the document.
@property (readonly) WordWdOMathBreakBin mathBinaryOperatorBreak;  // Gets or sets a value that specifies where Microsoft Word places binary operators when equations span two or more lines.
@property (readonly) WordWdOMathJc mathDefaultJustification;  // Gets or sets a value that indicates the default justification.
@property (copy, readonly) NSString *mathFontName;  // Gets or sets the name of the font that is used in a document to display equations.
@property (readonly) double mathLeftMargin;  // Gets or sets a value that specifies the left margin for equations.
@property (readonly) double mathRightMargin;  // Gets or sets a value that specifies the right margin for equations.
@property (readonly) WordWdOMathBreakSub mathSubtractionOperator;  // Gets or sets a value that specifies how Microsoft Word handles a subtraction operator that falls before a line break.
@property (readonly) double mathWrap;  // Gets or sets a value that specifies the placement of the second line of an equation that wraps to a new line.
@property (copy, readonly) NSString *name;  // Returns the name of the document.
@property (readonly) BOOL naryUsesSubscripts;  // Gets or sets a value that specifies the default location of limits for n-ary objects other than integrals
@property (copy) NSString *noLineBreakAfter;  // Returns or sets the kinsoku characters after which Microsoft Word will not break a line.
@property (copy) NSString *noLineBreakBefore;  // Returns or sets the kinsoku characters before which Microsoft Word will not break a line.
@property (copy) WordPageSetup *pageSetup;  // Returns or sets the page setup object.
@property (copy) NSString *password;  // Sets a password that must be supplied to open the specified document. This is write-only property
@property (copy, readonly) NSString *path;  // Returns the path to the document in HFS path style.
@property (copy, readonly) NSString *posixFullName;  // Returns the full name of the document in POSIX path style.
@property BOOL printFormsData;  // Returns or sets if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form.
@property BOOL printPostScriptOverText;  // Returns or sets if PRINT field instructions such as PostScript commands in a document are to be printed on top of text and graphics when a PostScript printer is used.
@property BOOL printRevisions;  // Returns or sets if revision marks are printed with the document. False if revision marks aren't printed that is, tracked changes are printed as if they'd been accepted.
@property (readonly) WordWdProtectionType protectionType;  // Returns the protection type for the specified document.
@property (readonly) BOOL readOnly;  // True if changes to the document cannot be saved to the original document.
@property BOOL readOnlyRecommended;  // Returns or set if Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only.
@property BOOL removeDateAndTime;  // Returns or sets if Microsoft Word removes date and time from revisions upon saving a document.
@property BOOL removePersonalInformation;  // Returns or sets if Microsoft Word removes all user information from comments, revisions, and the properties dialog box upon saving a document.
@property (readonly) WordWdSaveFormat saveFormat;  // Returns the file format of the specified document or file converter. Will be a unique number that specifies an external file converter or a constant.
@property BOOL saveFormsData;  // Returns or sets if Microsoft Word saves the data entered in a form as a tab-delimited record for use in a database.
@property BOOL saveSubsetFonts;  // Returns or sets if Microsoft Word saves a subset of the embedded TrueType fonts with the document.
@property BOOL saveVersionsOnClose;  // Sets or returns whether or not versions are automatically saved when a document is closed.
@property BOOL saved;  // Returns or set the saved state. True if the specified document or template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed.
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showGrammaticalErrors;  // Returns or sets if grammatical errors are marked by a wavy green line in the document. To view grammatical errors in your document, you must set the check grammar as you type property of the Word options class to true.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are shown.
@property BOOL showRevisions;  // Returns or sets if tracked changes in the specified document are shown on the screen.
@property BOOL showSpellingErrors;  // Returns or sets if Microsoft Word underlines spelling errors in the document.  To view spelling errors in your document, you must set the check grammar as you type property of the Word options class to true.
@property BOOL snapToGrid;  // Returns or sets if shape object or east asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in the specified document.
@property BOOL snapToShapes;  // Returns or sets if Microsoft Word automatically aligns shape object or east asian characters with invisible gridlines that go through the vertical and horizontal edges of other shape object or east asian characters in the document.
@property BOOL spellingChecked;  // True if a spelling check has been run on the document. False if some of the document hasn't been checked for spelling. To see if the document contains spelling errors, get the count of spelling errors for the document.
@property BOOL subdocumentsExpanded;  // Returns or set if the subdocuments in the document are expanded.
@property (copy, readonly) WordTextRange *textObject;  // Returns a range object that represents the main document story.
@property BOOL trackRevisions;  // Returns or sets if changes are tracked in the document.
@property (copy, readonly) NSArray<NSString *> *unavailableFonts;  // Returns a list of fonts used in the document that are not available on the current system.
@property BOOL updateStylesOnOpen;  // Returns or sets if the styles in the specified document are updated to match the styles in the attached template each time the document is opened.
@property (readonly) BOOL useDefaultMathSettings;  // Gets or sets a value that indicates whether to use the default math settings when creating new equations.
@property (readonly) BOOL useSmallFractions;  // Gets or sets a value that indicates whether to use small fractions in equations contained within the document.
@property (copy, readonly) WordWebOptions *webOptions;  // Returns the web options object.
@property (copy) NSString *writePassword;  // Sets a password for saving changes to the specified document. This is write-only property
@property (readonly) BOOL writeReserved;  // True if the specified document is protected with a write password.

@end

// Represents a dropped capital letter at the beginning of a paragraph.
@interface WordDropCap : SBObject <WordGenericMethods>

@property double distanceFromText;  // Returns or sets the distance in points between the dropped capital letter and the paragraph text. 
@property WordWdDropPosition dropPosition;  // Returns or sets the position of a dropped capital letter.
@property (copy) NSString *fontName;  // Returns or sets the name of the font for the dropped capital letter.
@property NSInteger linesToDrop;  // Returns or sets the height in lines of the specified dropped capital letter.

- (void) enable;  // Formats the first character in the specified paragraph as a dropped capital letter.

@end

// Represents a drop-down form field that contains a list of items in a form.
@interface WordDropDown : SBObject <WordGenericMethods>

- (SBElementArray<WordListEntry *> *) listEntries;

@property NSInteger dropDownDefault;  // Returns or sets the default drop-down item. The first item in a drop-down form field is 1, the second item is 2, and so on.
@property NSInteger dropDownValue;  // Returns or sets the number of the selected item in a drop-down form field.
@property (readonly) BOOL valid;  // Returns if the drop down object is valid.


@end

// A representation of the options associated with endnotes.
@interface WordEndnoteOptions : SBObject <WordGenericMethods>

@property (copy, readonly) WordTextRange *endnoteContinuationNotice;  // Returns a text range object that represents the endnote continuation notice.
@property (copy, readonly) WordTextRange *endnoteContinuationSeparator;  // Returns a text range object that represents the endnote continuation separator.
@property WordWdEndnoteLocation endnoteLocation;  // Returns or sets the position of all endnotes.
@property WordWdNoteNumberStyle endnoteNumberStyle;  // Returns or sets the number style for endnotes.
@property WordWdNumberingRule endnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks.
@property (copy, readonly) WordTextRange *endnoteSeparator;  // Returns a text range object that represents the endnote separator.
@property NSInteger endnoteStartingNumber;  // Returns or sets the starting endnote number.

- (void) endnoteConvert;  // Converts endnotes to footnotes, or vice versa.
- (void) swapWithFootnotes;  // Converts all footnotes in a document to endnotes and vice versa.

@end

// Represents an endnote.
@interface WordEndnote : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *noteReference;  // Returns a text range object that represents a endnote mark.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the endnote object.


@end

// Represents an envelope.
@interface WordEnvelope : SBObject <WordGenericMethods>

@property (copy, readonly) WordTextRange *address;  // Returns the envelope delivery address as a text range object.
@property double addressFromLeft;  // Returns or sets the distance in points between the left edge of the envelope and the delivery address.
@property double addressFromTop;  // Returns or sets the distance in points between the top edge of the envelope and the delivery address.
@property (copy, readonly) WordWordStyle *addressStyle;  // Returns a Word style object that represents the delivery address style for the envelope.
@property BOOL defaultFaceUp;  // Returns or sets if envelopes are fed face up by default.
@property double defaultHeight;  // Returns or sets the default envelope height, in points.
@property BOOL defaultOmitReturnAddress;  // Returns or sets if the return address is omitted from envelopes by default.
@property WordWdEnvelopeOrientation defaultOrientation;  // Returns or sets the default orientation for feeding envelopes.
@property BOOL defaultPrintFIMA;  // Returns or sets if a Facing Identification Mark FIM-A to envelopes by default. A FIM-A code is used to presort courtesy reply mail. The default print barcode property must be set to true before this property is set.
@property BOOL defaultPrintBarCode;  // Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set
@property (copy) NSString *defaultSize;  // Returns or sets the default envelope size. If you set either the default height or default width property, the property  is automatically changed to return custom size.
@property double defaultWidth;  // Returns or sets the default envelope width, in points.
@property WordWdPaperTray feedSource;  // Returns or sets the paper tray for the envelope.
@property (copy, readonly) WordTextRange *returnAddress;  // Returns the envelope return address as a text range object.
@property double returnAddressFromLeft;  // Returns or sets the distance in points between the left edge of the envelope and the return address.
@property double returnAddressFromTop;  // Returns or sets the distance in points between the top edge of the envelope and the return address.
@property (copy, readonly) WordWordStyle *returnAddressStyle;  // Returns a Word style object that represents the return address style for the envelope.

- (void) insertEnvelopeDataExtractAddress:(BOOL)extractAddress address:(NSString *)address autoText:(NSString *)autoText omitReturnAddress:(BOOL)omitReturnAddress returnAddress:(NSString *)returnAddress returnAutotext:(NSString *)returnAutotext printBarCode:(BOOL)printBarCode printFIMA:(BOOL)printFIMA envelopeSize:(NSString *)envelopeSize envelopeHeight:(NSInteger)envelopeHeight envlopeWidth:(NSInteger)envlopeWidth feedSource:(BOOL)feedSource addressFromLeft:(NSInteger)addressFromLeft addressFromTop:(NSInteger)addressFromTop returnAddressFromLeft:(NSInteger)returnAddressFromLeft returnAddressFromTop:(NSInteger)returnAddressFromTop defaultFaceUp:(BOOL)defaultFaceUp defaultOrientation:(WordWdEnvelopeOrientation)defaultOrientation sizeFromPageSetup:(BOOL)sizeFromPageSetup showPageSetupDialog:(BOOL)showPageSetupDialog createNewDocument:(BOOL)createNewDocument;  // Inserts an envelope as a separate section at the beginning of the specified document.
- (void) printOutEnvelopeExtractAddress:(BOOL)extractAddress address:(NSString *)address autoText:(NSString *)autoText omitReturnAddress:(BOOL)omitReturnAddress returnAddress:(NSString *)returnAddress returnAutotext:(NSString *)returnAutotext printBarCode:(BOOL)printBarCode printFIMA:(BOOL)printFIMA envelopeSize:(NSString *)envelopeSize envelopeHeight:(NSInteger)envelopeHeight envlopeWidth:(NSInteger)envlopeWidth feedSource:(BOOL)feedSource addressFromLeft:(NSInteger)addressFromLeft addressFromTop:(NSInteger)addressFromTop returnAddressFromLeft:(NSInteger)returnAddressFromLeft returnAddressFromTop:(NSInteger)returnAddressFromTop defaultFaceUp:(BOOL)defaultFaceUp defaultOrientation:(WordWdEnvelopeOrientation)defaultOrientation sizeFromPageSetup:(BOOL)sizeFromPageSetup showPageSetupDialog:(BOOL)showPageSetupDialog;
- (void) updateDocument;  // Updates the envelope in the document with the current envelope settings.

@end

@interface WordFieldOptions : SBObject <WordGenericMethods>

@property BOOL locked;


@end

// Represents a field.
@interface WordField : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) WordTextRange *fieldCode;  // Returns a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters.
@property (readonly) WordWdFieldKind fieldKind;  // Returns the type of link for a field object.
@property (copy) NSString *fieldText;  // Returns or sets data in an ADDIN field. The data is not visible in the field code or result. It is only accessible by returning the value of the data property. If the field isn't an ADDIN field, this property will cause an error.
@property (readonly) WordWdFieldType fieldType;  // Returns the field type.
@property (copy, readonly) WordInlineShape *inlineShape;  // Returns the inline shape object associated with this field
@property (copy, readonly) WordLinkFormat *linkFormat;  // Returns the link format object associated with this field object.
@property BOOL locked;  // Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results.
@property (copy, readonly) WordField *nextField;  // Returns the next field object.
@property (copy, readonly) WordField *previousField;  // Returns the previous field object.
@property (copy) WordTextRange *resultRange;  // Returns or sets a text range object that represents a field's result. You can access a field result without changing the view from field codes.
@property BOOL showCodes;  // Returns or sets if field codes are displayed for the specified field instead of field results.

- (void) clickObject;  // Clicks the specified field. If the field is a GOTOBUTTON field, this method moves the insertion point to the specified location or selects the specified bookmark. If the field is a HYPERLINK field, this method jumps to the target location.
- (BOOL) updateField;  // Updates the result of the field object. When applied to a field object, returns true if the field is updated successfully.

@end

// Represents a file converter that's used to open or save files.
@interface WordFileConverter : SBObject <WordGenericMethods>

@property (readonly) BOOL canOpen;  // Returns true if the specified file converter is designed to open files.
@property (readonly) BOOL canSave;  // Returns true if the specified file converter is designed to save files.
@property (copy, readonly) NSString *className;  // Returns a unique name that identifies the file converter.
@property (copy, readonly) NSString *extensions;  // Returns the file name extensions associated with the specified file converter object.
@property (copy, readonly) NSString *formatName;  // Returns the name of the specified file converter.
@property (copy, readonly) NSString *name;  // Returns the name of the file converter.
@property (readonly) NSInteger openFormat;  // Returns the file format of the specified file converter. It will be a unique number that represents an external file converter.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified file converter in HFS path style. 
@property (readonly) NSInteger saveFormat;  // Returns the file format of the specified document or file converter. It will be a unique number that specifies an external file converter.


@end

// Represents the criteria for a find operation.
@interface WordFind : SBObject <WordGenericMethods>

@property (copy) NSString *content;  // Returns or sets the text in the find object.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this find object.
@property BOOL format;  // Returns or set if formatting is included in the find operation.
@property BOOL forward;  // Returns or sets if the find operation searches forward through the document. False if it searches backward through the document.
@property (readonly) BOOL found;  // True if the search produces a match.
@property (copy, readonly) WordFrame *frame;  // Returns the frame object associated with the find object.
@property NSInteger highlight;  // Returns or sets if highlight formatting is included in the find criteria
@property WordWdLanguageID languageID;  // Returns or sets the language for the find object
@property WordWdLanguageID languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property BOOL matchAllWordForms;  // Returns or sets if all forms of the text to find are found by the find operation for instance, if the text to find is sit, sat and sitting are found as well.
@property BOOL matchByte;  // Returns or sets if Microsoft Word distinguishes between full-width and half-width letters or characters during a search.
@property BOOL matchCase;  // Returns or sets if the find operation is case sensitive.
@property BOOL matchFuzzy;  // Returns or sets if Microsoft Word uses the nonspecific search options for Japanese text during a search.
@property BOOL matchSoundsLike;  // Returns or sets if words that sound similar to the text to find are returned by the find operation.
@property BOOL matchWholeWord;  // Returns or sets if the find operation locates only entire words and not text that's part of a larger word.
@property BOOL matchWildcards;  // Returns or sets if the text to find contains wildcards.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with the find object.
@property (copy, readonly) WordReplacement *replacement;  // Returns the replacement object associated with the find object.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style associated with the find object.
@property WordWdLanguageID supplementalLanguageID;  // Returns or sets the language for the text range object
@property WordWdFindWrap wrap;  // Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached or vice versa if forward is set to false or if the search text isn't found in the specified selection or range. 

- (void) clearAllFuzzyOptions;  // Clears all nonspecific search options associated with Japanese text.
- (WordWdInsertionPoint) executeFindFindText:(NSString *)findText matchCase:(BOOL)matchCase matchWholeWord:(BOOL)matchWholeWord matchWildcards:(BOOL)matchWildcards matchSoundsLike:(BOOL)matchSoundsLike matchAllWordForms:(BOOL)matchAllWordForms matchForward:(BOOL)matchForward wrapFind:(WordWdFindWrap)wrapFind findFormat:(BOOL)findFormat replaceWith:(NSString *)replaceWith replace:(WordWdReplace)replace;  // Runs the specified find operation. Returns true if the find operation is successful.
- (void) setAllFuzzyOptions;  // Activates all nonspecific search options associated with Japanese text.

@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface WordFont : SBObject <WordGenericMethods>

@property BOOL allCaps;  // Returns or sets if the font is formatted as all capital letters.
@property WordWdAnimation animation;  // Returns or sets the type of animation applied to the font.
@property (copy) NSString *asciiName;  // Returns or sets the font used for Latin text characters with character codes from 0 through 127. 
@property BOOL bold;  // Returns or sets if the font is formatted as bold.
@property BOOL boldBi;  // Returns or sets if the font is formatted as bold.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the RGB color of the font object.
@property WordWdColorIndex colorIndex;  // Returns or sets the color for the font object using an index.
@property WordMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified border or font object.
@property (copy) NSString *complexScriptName;  // Returns or sets a bidi name - complex script name.
@property BOOL disableCharacterSpaceGrid;  // Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding font object.
@property BOOL doubleStrikeThrough;  // Returns or sets if the specified font is formatted as double strikethrough text.
@property (copy) NSString *eastAsianName;  // Returns or sets an East Asian font name.
@property BOOL emboss;  // Returns or sets if the specified font is formatted as embossed.
@property WordWdEmphasisMark emphasisMark;  // Returns or sets the emphasis mark for a character or designated character string.
@property BOOL engrave;  // Returns or sets if the specified font is formatted as engraved.
@property NSInteger fontPosition;  // Returns or sets the position of text in points relative to the base line. A positive number raises the text, and a negative number lowers it.
@property double fontSize;  // Returns or sets the font size.
@property BOOL hidden;  // Returns or sets if the font is formatted as hidden text.
@property BOOL italic;  // Returns or sets if the font is formatted as italic.
@property BOOL italicBi;  // Returns or sets if the font is formatted as italic.
@property double kerning;  // Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically.
@property (copy) NSString *name;  // Returns or sets the font name associated with this font object.
@property (copy) NSString *otherName;  // Returns or sets the font used for characters with character codes from 128 through 255.
@property BOOL outline;  // Returns or sets if the specified font is formatted as outline.
@property NSInteger scaling;  // Returns or sets the scaling percentage applied to the font. This property stretches or compresses text horizontally as a percentage of the current size. The scaling range is from 1 through 600.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the font object.
@property BOOL shadow;  // Returns or sets if the specified font is formatted as shadowed.
@property BOOL smallCaps;  // Returns or sets if the font is formatted as small capital letters.
@property double spacing;  // Returns or sets the spacing in points between characters.
@property BOOL strikeThrough;  // Returns or sets if the font is formatted as strikethrough text.
@property BOOL subscript;  // Returns or sets if the font is formatted as subscript.
@property BOOL superscript;  // Returns or sets if the font is formatted as superscript.
@property WordWdUnderline underline;  // Returns or sets the type of underline applied to the font.
@property (copy) NSArray<NSNumber *> *underlineColor;  // Returns or sets the RGB color of the underline for the font object.
@property WordMsoThemeColorIndex underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.

- (void) growFont;  // Increases the font size to the next available size. If the selection or range contains more than one font size, each size is increased to the next available setting.
- (void) reset;  // Removes changes that were made to a font.
- (void) setAsFontTemplateDefault;  // Sets the font formatting as the default for the active document and all new documents based on the active template. The default font formatting is stored in the Normal style.
- (void) shrinkFont;  // Decreases the font size to the next available size. If the selection or range contains more than one font size, each size is decreased to the next available setting.

@end

// A representation of the options associated with footnotes.
@interface WordFootnoteOptions : SBObject <WordGenericMethods>

@property (copy, readonly) WordTextRange *footnoteContinuationNotice;  // Returns a text range object that represents the footnote continuation notice.
@property (copy, readonly) WordTextRange *footnoteContinuationSeparator;  // Returns a text range object that represents the footnote continuation separator.
@property WordWdFootnoteLocation footnoteLocation;  // Returns or sets the position of all footnotes.
@property WordWdNoteNumberStyle footnoteNumberStyle;  // Returns or sets the number style for footnotes.
@property WordWdNumberingRule footnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. 
@property (copy, readonly) WordTextRange *footnoteSeparator;  // Returns a text range object that represents the footnote separator.
@property NSInteger footnoteStartingNumber;  // Returns or sets the starting footnote number. 

- (void) footnoteConvert;  // Converts endnotes to footnotes, or vice versa.
- (void) swapWithEndnotes;  // Converts all footnotes in a document to endnotes and vice versa.

@end

// Represents a footnote positioned at the bottom of the page or beneath text.
@interface WordFootnote : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *noteReference;  // Returns a text range object that represents a footnote mark.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the footnote object.


@end

// Represents a single form field.
@interface WordFormField : SBObject <WordGenericMethods>

@property BOOL calculateOnExit;  // Returns or sets if references to the specified form field are automatically updated whenever the field is exited.
@property (copy, readonly) WordCheckBox *checkBox;  // Returns the check box object associated with the form filed object.
@property (copy, readonly) WordDropDown *dropDown;  // Returns the drop down object associated with the form filed object.
@property BOOL enabled;  // Returns or sets if this form field object is enabled.
@property (copy) NSString *formFieldResult;  // Returns or sets a string that represents the result of the specified form field.
@property (readonly) WordWdFieldType formFieldType;  // The type of this form field.
@property (copy) NSString *helpText;  // Returns or sets the text that's displayed in a message box when the form field has the focus and the user presses F1.
@property (copy) NSString *name;  // Returns or sets the name of the form field.
@property (copy, readonly) WordFormField *nextFormField;  // Returns the next form field object.
@property BOOL ownHelp;  // Returns or set the source of the text that's displayed in a message box when a form field has the focus and the user presses F1. If true, the text specified by the helptext property is displayed. If false, the text in the autotext entry is displayed
@property BOOL ownStatus;  // Returns or sets the source of the text that's displayed in the status bar when a form field has the focus. If true, the text specified by the status text property is displayed. If false, the text of the autotext entry is displayed.
@property (copy, readonly) WordFormField *previousFormField;  // Returns the previous form field object.
@property (copy) NSString *statusText;  // Returns or sets the text that's displayed in the status bar when a form field has the focus.
@property (copy, readonly) WordTextInput *textInput;  // Returns the text input object associated with the form filed object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the form field object.


@end

// Represents a frame.
@interface WordFrame : SBObject <WordGenericMethods>

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this frame object.
@property double height;  // Returns or sets the height of the object.
@property WordWdFrameSizeRule heightRule;  // Returns or sets the rule for determining the height of the specified frame.
@property double horizontalDistanceFromText;  // Returns or sets the horizontal distance between a frame and the surrounding text, in points.
@property double horizontalPosition;  // Returns or sets the horizontal distance between the edge of the frame and the item specified by the relative horizontal position property.
@property BOOL lockAnchor;  // Returns or sets if the specified frame is locked. The frame anchor indicates where the frame will appear in Draft view. You cannot reposition a locked frame anchor.
@property WordWdRelativeHorizontalPosition relativeHorizontalPosition;  // Returns or sets what the horizontal position of a frame is relative.
@property WordWdRelativeVerticalPosition relativeVerticalPosition;  // Returns or sets what the vertical position of a frame is relative.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the frame object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the frame object.
@property BOOL textWrap;  // Returns or sets if document text wraps around the specified frame.
@property double verticalDistanceFromText;  // Returns or sets the vertical distance in points between a frame and the surrounding text.
@property double verticalPosition;  // Returns or sets the vertical distance between the edge of the frame and the item specified by the relative vertical position property. 
@property double width;  // Returns or sets the width of the object.
@property WordWdFrameSizeRule widthRule;  // Returns or sets the rule used to determine the width of a frame.


@end

// Represents a single header or footer.
@interface WordHeaderFooter : SBObject <WordGenericMethods>

- (SBElementArray<WordPageNumber *> *) pageNumbers;
- (SBElementArray<WordShape *> *) shapes;

@property (readonly) WordWdHeaderFooterIndex headerFooterIndex;  // Returns a constant that represents the specified header or footer in a document or section.
@property (readonly) BOOL isHeader;  // Returns true if this object is a header.
@property BOOL linkToPrevious;  // Returns or sets if the specified header or footer is linked to the corresponding header or footer in the previous section. When a header or footer is linked, its contents are the same as in the previous header or footer.
@property (copy, readonly) WordPageNumberOptions *pageNumberOptions;  // Return the page number options object associated with this header footer object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the header or footer.


@end

// Represents a style used to build a table of contents or figures.
@interface WordHeadingStyle : SBObject <WordGenericMethods>

@property NSInteger level;  // Returns or sets the level for the heading style in a table of contents or table of figures.
@property WordWdBuiltinStyle style;  // Returns or sets the style associated with the heading style object.


@end

// Represents a hyperlink.
@interface WordHyperlinkObject : SBObject <WordGenericMethods>

@property (copy) NSString *emailSubject;  // Returns or sets the text string for the specified hyperlinkÔøïs subject line. The subject line is appended to the hyperlinkÔøïs Internet address, or URL.
@property (readonly) BOOL extraInfoRequired;  // Returns true if extra information is required to resolve the specified hyperlink.
@property (copy) NSString *hyperlinkAddress;  // Returns or sets the address, for example, a file name or URL of the specified hyperlink. 
@property (readonly) WordMsoHyperlinkType hyperlinkType;  // The type of this hyperlink
@property (copy, readonly) NSString *name;  // The name of this hyperlink object.
@property (copy) NSString *screenTip;  // Returns or sets the text that appears as a screen tip when the mouse pointer is positioned over the specified hyperlink.
@property (copy, readonly) WordShape *shape;
@property (copy) NSString *subAddress;  // Returns or sets a named location in the destination of the specified hyperlink. 
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the hyperlink.
@property (copy) NSString *textToDisplay;  // Returns or sets the specified hyperlink's visible text in a document.

- (void) createNewDocumentForHyperlinkFileName:(NSString *)fileName editNow:(BOOL)editNow overwrite:(BOOL)overwrite;  // Creates a new document linked to the specified hyperlink.
- (void) followNewWindow:(BOOL)newWindow extraInfo:(NSString *)extraInfo;  // Displays a cached document associated with the specified hyperlink object, if it's already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

@end

// Represents a single index.
@interface WordIndex : SBObject <WordGenericMethods>

@property BOOL accentedLetters;  // Returns or sets if the specified index contains separate headings for accented letters, for example, words that begin with Ôøã are under one heading and words that begin with A are under another.
@property WordWdHeadingSeparator headingSeparator;  // Returns or sets the text between alphabetic groups, entries that start with the same letter in the index.  
@property WordWdIndexFilter indexFilter;  // Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index. 
@property WordWdIndexType indexType;  // Returns or sets the index type.
@property NSInteger numberOfColumns;  // Sets or returns the number of columns for each page of an index.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in an index. 
@property WordWdIndexSortBy sortBy;  // Returns or sets the sorting criteria for the specified index.
@property WordWdTabLeader tabLeader;  // Returns or sets the character between entries and their page numbers in an index. 
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the index


@end

// Represents a custom key assignment in the current context.
@interface WordKeyBinding : SBObject <WordGenericMethods>

@property (copy, readonly) SBObject *bindingContext;  // Returns an object that represents the storage location of the specified key binding. This property can return a document or template object.
@property (copy, readonly) NSString *bindingKeyString;  // Returns the key combination string for the specified keys.
@property (copy, readonly) NSString *command;  // Returns the command assigned to the specified key combination.
@property (copy, readonly) NSString *commandParameter;  // Returns the command parameter assigned to the specified shortcut key.
@property (readonly) WordWdKeyCategory keyCategory;  // Returns the type of item assigned to the specified key binding.
@property (readonly) NSInteger keyCode;  // Returns a unique number for the first key in the specified key binding.  You create this number by using the build key code method.
@property (readonly) NSInteger key_code_2;  // Returns a unique number for the second key in the specified key binding. You create this number by using the build key code method.
- (BOOL) protected;  // Returns true if you cannot change the specified key binding in the customize keyboard dialog box. 

- (void) disable;  // Removes the specified key combination if it's currently assigned to a command. After you use this method, the key combination has no effect.
- (void) executeKeyBinding;  // Runs the command associated with the specified key combination.
- (void) rebindKeyCategory:(WordWdKeyCategory)keyCategory command:(NSString *)command commandParameter:(NSString *)commandParameter;  // Changes the command assigned to the specified key binding.

@end

// Represents the elements of a letter created by the letter wizard.
@interface WordLetterContent : SBObject <WordGenericMethods>

@property (copy) NSString *attentionLine;  // Returns or sets the attention line text for a letter created by the letter wizard.
@property (copy) NSString *ccList;  // Returns or sets the carbon copy recipients for a letter created by the letter wizard.
@property (copy) NSString *closing;  // Returns or sets the closing text for a letter created by the letter wizard, for example, Sincerely yours.
@property (copy) NSString *dateFormat;  // Returns or sets the date for a letter created by the letter wizard. 
@property NSInteger enclosureCount;  // Returns or sets the number of enclosures for a letter created by the letter wizard. 
@property BOOL includeHeaderFooter;  // Returns or sets if the header and footer from the page design template are included in a letter created by the letter wizard.
@property WordWdLetterStyle letterStyle;  // Returns or sets the layout of a letter created by the letter wizard.
@property BOOL letterhead;  // Returns or sets if space is reserved for a preprinted letterhead in a letter created by the letter wizard.
@property WordWdLetterheadLocation letterheadLocation;  // Returns or sets the location of the preprinted letterhead in a letter created by the letter wizard.
@property double letterheadSize;  // Returns or sets the amount of space in points to be reserved for a preprinted letterhead in a letter created by the letter wizard.
@property (copy) NSString *mailingInstructions;  // Returns or sets the mailing instruction text for a letter created by the letter wizard, for example, Certified Mail.
@property (copy) NSString *pageDesign;  // Returns or sets the name of the template attached to the document created by the letter wizard.
@property (copy) NSString *recipientAddress;  // Returns or sets the return address for a letter created with the letter wizard.
@property (copy) NSString *recipientName;  // Returns or sets the name of the person who'll be receiving the letter created by the letter wizard.
@property (copy) NSString *recipientReference;  // Returns or sets the reference line, for example, In reply to: for a letter created by the letter wizard.
@property (copy) NSString *returnAddress;  // Returns or sets the return address for a letter created with the letter wizard. 
@property (copy) NSString *salutation;  // Returns or sets the salutation text for a letter created by the letter wizard.
@property WordWdSalutationType salutationType;  // Returns or sets the type of salutation for a letter created by the letter wizard.  
@property (copy, readonly) NSString *senderCity;
@property (copy) NSString *senderCompany;  // Returns or sets the company name of the person creating a letter with the letter wizard.
@property (copy) NSString *senderInitials;  // Returns or sets the initials of the person creating a letter with the letter wizard. 
@property (copy) NSString *senderJobTitle;  // Returns or sets the job title of the person creating a letter with the letter wizard.
@property (copy) NSString *senderName;  // Returns or sets the name of the person creating a letter with the letter wizard. 
@property (copy) NSString *subject;  // Returns or sets the subject text of a letter created by the letter wizard.


@end

// Represents line numbers in the left margin or to the left of each newspaper-style column.
@interface WordLineNumbering : SBObject <WordGenericMethods>

@property BOOL activeLine;  // Returns or sets if line numbering is active for the specified document, section, or sections.
@property NSInteger countBy;  // Returns or sets the numeric increment for line numbers. For example, if the count by property is set to 5, every fifth line will display the line number. Line numbers are only displayed in print layout view and print preview.
@property double distanceFromText;  // Returns or sets the distance in points between the right edge of line numbers and the left edge of the document text.
@property WordWdNumberingRule restartMode;  // Returns or sets the way line numbering runsÔøë that is, whether it starts over at the beginning of a new page or section or runs continuously.
@property NSInteger startingNumber;  // Returns or sets the starting line number.


@end

// Represents the linking characteristics for a picture.
@interface WordLinkFormat : SBObject <WordGenericMethods>

@property BOOL autoUpdate;  // Returns or sets if the specified link is updated automatically when the container file is opened or when the source file is changed.
@property (readonly) WordWdLinkType linkType;  // Returns the link type.
@property BOOL locked;  // Returns or sets if inline shape object is locked to prevent automatic updating.
@property BOOL savePictureWithDocument;  // Returns or sets if the specified picture is saved with the document.
@property (copy) NSString *sourceFullName;  // Returns or sets the path and name of the source file for the specified picture.
@property (copy, readonly) NSString *sourceName;  // Returns the name of the source file for the specified picture.
@property (copy, readonly) NSString *sourcePath;  // Returns the path of the source file for the specified picture.

- (void) breakLink;  // Breaks the link between the source file and the specified picture.

@end

// Represents an item in a drop-down form field.
@interface WordListEntry : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of this list entry object.


@end

// Represents the list formatting attributes that can be applied to the paragraphs in a range.
@interface WordListFormat : SBObject <WordGenericMethods>

@property (copy, readonly) WordWordList *WordList;  // Returns a Word list object that represents the first formatted list contained in the list format object.
@property NSInteger listLevelNumber;  // Returns or sets the list level for the first paragraph in the list format object.
@property (copy) WordInlineShape *listPictureBullet;
@property (copy, readonly) NSString *listString;  // Returns a string that represents the appearance of the list value of the first paragraph in the range for the list format object. For example, the second paragraph in an alphabetic list would return B.
@property (copy, readonly) WordListTemplate *listTemplate;  // Returns a list template object that represents the list formatting for the list format object.
@property (readonly) WordWdListType listType;  // Returns the type of lists that are contained in the range for the list format object.
@property (readonly) NSInteger listValue;  // Returns the numeric value of the first paragraph in the range for the specified list format object. For example, the list value property applied to the second paragraph in an alphabetic list would return 2.
@property (readonly) BOOL singleList;  // Returns if the specified list format object contains only one list.
@property (readonly) BOOL singleListTemplate;  // True if the entire list format object uses the same list template

- (void) applyBulletDefaultDefaultListBehavior:(WordWdDefaultListBehavior)defaultListBehavior;  // Adds bullets and formatting to the paragraphs in the range for the list format object. If the paragraphs are already formatted with bullets, this method removes the bullets and formatting.
- (void) applyListFormatTemplateListTemplate:(WordListTemplate *)listTemplate continuePreviousList:(BOOL)continuePreviousList applyTo:(WordWdListApplyTo)applyTo defaultListBehavior:(WordWdDefaultListBehavior)defaultListBehavior;  // Applies a set of list-formatting characteristics to the list format object
- (void) applyNumberDefaultDefaultListBehavior:(WordWdDefaultListBehavior)defaultListBehavior;  // Adds the default numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.
- (void) applyOutlineNumberDefaultDefaultListBehavior:(WordWdDefaultListBehavior)defaultListBehavior;  // Adds the default outline-numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as an outline-numbered list, this method removes the numbers and formatting.
- (void) listIndent;  // Increases the list level of the paragraphs in the range for the list format object, in increments of one level.
- (void) listOutdent;  // Decreases the list level of the paragraphs in the range for the list format object, in increments of one level.

@end

// Represents a single gallery of list formats.
@interface WordListGallery : SBObject <WordGenericMethods>

- (SBElementArray<WordListTemplate *> *) listTemplates;

- (BOOL) modifiedIndex:(NSInteger)index;  // if the specified list template is not the built-in list template for that position in the list gallery. Index goes from 1 to 7
- (void) resetListGalleryIndex:(NSInteger)index;  // Resets the list template specified by index for the specified list gallery to the built-in list template format.

@end

// Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list.
@interface WordListLevel : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this list level object.
@property (copy) NSString *linkedStyle;  // Returns or sets the name of the style that's linked to the specified list level object.
@property WordWdListLevelAlignment listLevelAlignment;  // Returns or sets a constant that represents the alignment for the list level of the list template.
@property (copy) NSString *numberFormat;  // Returns or sets the number format for the specified list level.
@property double numberPosition;  // Returns or sets the position in points of the number or bullet for the specified list level object.
@property WordWdListNumberStyle numberStyle;  // Returns or sets the number style for the list level object.
@property (copy, readonly) WordInlineShape *pictureBullet;
@property NSInteger resetOnHigher;  // Returns or sets the list level that must appear before the specified list level restarts numbering at 1.
@property NSInteger startAt;  // Returns or sets the starting number for the specified list level object.
@property double tabPosition;  // Returns or sets the tab position for the specified list level object.
@property double textPosition;  // Returns or sets the position in points for the second line of wrapping text for the specified list level object.
@property WordWdTrailingCharacter trailingCharacter;  // Returns or sets the character inserted after the number for the specified list level.

- (WordInlineShape *) applyPictureBulletPath:(NSString *)path;

@end

// Represents a single list template that includes all the formatting that defines a list.
@interface WordListTemplate : SBObject <WordGenericMethods>

- (SBElementArray<WordListLevel *> *) listLevels;

@property (copy) NSString *name;  // Returns or sets the name of this list template object.
@property BOOL outlineNumbered;  // Returns or sets if the specified list template object is outline numbered.

- (WordListTemplate *) convertLevel:(NSInteger)level;  // Converts a multiple-level list to a single-level list, or vice versa.

@end

// Represents a mailing label.
@interface WordMailingLabel : SBObject <WordGenericMethods>

- (SBElementArray<WordCustomLabel *> *) customLabels;

@property (copy) NSString *defaultLabelName;  // Returns or sets the name for the default mailing label.
@property WordWdPaperTray defaultLaserTray;  // Returns or sets the default paper tray that contains sheets of mailing labels
@property BOOL defaultPrintBarCode;  // Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set

- (WordDocument *) createNewMailingLabelDocumentName:(NSString *)name address:(NSString *)address autoText:(NSString *)autoText extractAddress:(BOOL)extractAddress laserTray:(WordWdPaperTray)laserTray singleLabel:(BOOL)singleLabel row:(NSInteger)row column:(NSInteger)column;  // Creates a new label document using either the default label options or ones that you specify. Returns a document object that represents the new document.
- (void) printOutMailingLabelName:(NSString *)name address:(NSString *)address extractAddress:(BOOL)extractAddress laserTray:(WordWdPaperTray)laserTray singleLabel:(BOOL)singleLabel row:(NSInteger)row column:(NSInteger)column;  // Prints a label or a page of labels with the same address.

@end

// Represents an equation that has an accent mark above the base.
@interface WordMathAccent : SBObject <WordGenericMethods>

- (NSInteger) char;  // Gets or sets an integer that represents the accent character for the accent object.
- (void) setChar: (NSInteger) character;
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.


@end

// Represents an individual entry in the Math AutoCorrect engine.
@interface WordMathAutocorrectEntry : SBObject <WordGenericMethods>

@property (copy) NSString *autocorrectValue;  // Gets or sets a string that represents the contents of an equation auto correct entry.
@property (readonly) NSInteger entry_index;  // Gets an integer that represents the position of an item in the collection.
@property (copy) NSString *name;  // Gets or sets a string that represents the name of an equation auto correct entry.


@end

// Represents the Math AutoCorrect feature in Microsoft Word.
@interface WordMathAutocorrect : SBObject <WordGenericMethods>

- (SBElementArray<WordMathAutocorrectEntry *> *) mathAutocorrectEntries;
- (SBElementArray<WordMathRecognizedFunction *> *) mathRecognizedFunctions;

@property BOOL replaceText;  // Gets or sets whether Microsoft Word automatically replaces strings in equations with the corresponding math AutoCorrect definitions.
@property BOOL useOutsideEquations;  // Gets or sets whether Microsoft Word uses math autocorrect rules outside equations in a document.

- (void) addMathAcEntryName:(NSString *)name value:(NSString *)value;  // Creates an equation auto correct entry.
- (void) addMathRecognizedFunctionName:(NSString *)name;  // Creates a new recognized function.

@end

// Represents the mathematical overbar for an object in an equation.
@interface WordMathBar : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isTopBar;  // Gets or sets the position of a bar in a bar object. True specifies a mathematical overbar. False specifies a mathematical underbar


@end

// Represents an invisible box around an equation or part of an equation to which you can assign properties that affect the layout or mathematical formatting of the entire box.
@interface WordMathBorderBox : SBObject <WordGenericMethods>

@property BOOL bottomLeftToTopRightStrikethrough;  // Gets or sets if a diagonal strikethrough from lower left to upper right is drawn.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hideBottom;  // Gets or sets whether to hide the bottom border of an equation's bounding box.
@property BOOL hideLeft;  // Gets or sets whether to hide the left border of an equation's bounding box.
@property BOOL hideRight;  // Gets or sets whether to hide the right border of an equation's bounding box.
@property BOOL hideTop;  // Gets or sets whether to hide the top border of an equation's bounding box.
@property BOOL horizontalStrikethrough;  // Gets or sets if a horizontal strikethrough is drawn.
@property BOOL topLeftToBottomRightStrikethrough;  // Gets or sets if a diagonal strikethrough from upper left to lower right is drawn.
@property BOOL verticalStrikethrough;  // Gets or sets if vertical strikethrough is drawn.


@end

// Represents an invisible box around an equation or part of an equation to which you can apply properties that affect the mathematical or formatting properties, such as line breaks.
@interface WordMathBox : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isDifferential;  // Gets or sets whether the box acts as the mathematical differential, in which case the box receives the appropriate horizontal spacing for a differential.
@property BOOL isOperatorEmulator;  // Gets or sets if the box and its contents behave as a single operator and inherits the properties of an operator.
@property BOOL nonBreaking;  // Gets or sets whether breaks are allowed inside the box object.


@end

// Represents individual line breaks in an equation.
@interface WordMathBreak : SBObject <WordGenericMethods>

@property NSInteger alignAt;  // Gets or sets an integer that represents the operator in one line, to which to align consecutive lines in an equation.
@property (copy, readonly) WordTextRange *textRange;  // Returns a text range object that represents the portion of a document that is contained in the specified object.


@end

// Represents a delimiter object, consisting of opening and closing delimiters, e.g. parentheses, braces, brackets, or vertical bars, and one or more elements contained inside the delimiters.
@interface WordMathDelimiter : SBObject <WordGenericMethods>

- (SBElementArray<WordMathObject *> *) mathObjects;

@property NSInteger beginningCharacter;  // Gets or sets an Integer that represents the beginning delimiter character in a delimiter object.
@property WordWdOMathShapeType delimiterShape;  // Gets or sets the appearance of delimiters, e.g. parentheses, braces, and brackets, in relationship to the content that they surround.
@property NSInteger endingCharacter;  // Gets or sets an Integer that represents the ending delimiter character in a delimiter object.
@property BOOL hideLeftDelimiter;  // Gets or sets whether to hide the opening delimiter in a delimiter object.
@property BOOL hideRightDelimiter;  // Gets or sets whether to hide the closing delimiter in a delimiter object.
@property NSInteger separatorCharacter;  // Gets or sets an Integer that represents the separator character in a delimiter object when the delimiter object contains two or more arguments.
@property BOOL shouldGrow;  // Gets or sets whether delimiter characters grow to the full height of the arguments that they contain.


@end

// Represents a mathematical equation array object, consisting of one or more equations that can be vertically justified as a unit respect to surrounding text on the line.
@interface WordMathEquationArray : SBObject <WordGenericMethods>

- (SBElementArray<WordMathObject *> *) mathObjects;

@property BOOL distributeEqually;  // Gets or sets if the equations in an equation array are distributed equally within the margins of its container, such as a column, cell, or page width.
@property NSInteger rowSpacing;  // Gets or sets an integer that represents the spacing between the rows in an equation array.
@property WordWdOMathSpacingRule rowSpacingRule;  // Gets or sets the spacing rule that defines spacing in an equation array.
@property BOOL useMaxWidth;  // Gets or sets whether the equations in an equation array are spaced to the maximum width of the equation array.
@property WordWdOMathVertAlignType verticalAlignment;  // Gets or sets the type of vertical alignment for an equation array with respect to the text that surrounds the array.


@end

// Represents a fraction, consisting of a numerator and denominator separated by a fraction bar. The fraction bar can be horizontal or diagonal, depending on the fraction properties.
@interface WordMathFraction : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *denominator;  // Returns a math object object that represents the denominator for an equation that contains a fraction.
@property WordWdOMathFracType fractionType;  // Gets or sets the layout of a fraction, whether it is stacked, skewed, linear, or without a fraction bar.
@property (copy, readonly) WordMathObject *numerator;  // Returns a math object object that represents the numerator for the fraction.


@end

// Represents a math func object that represents a type of mathematical function that consists of a function name, such as sin or cos, and an argument.
@interface WordMathFunc : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *funcName;  // Returns an OMath object that represents the name of a mathematical function, such as sin or cos.


@end

// Represents a mathematical function or structure such as fractions, integrals, sums, and radicals.
@interface WordMathFunction : SBObject <WordGenericMethods>

- (SBElementArray<WordMathObject *> *) mathObjects;

@property (copy, readonly) WordMathAccent *accent;  // Returns a math accent object that represents a base character with a combining accent mark.
@property (copy, readonly) WordMathBorderBox *borderBox;  // Returns a math border box object that represents a border drawn around an equation or part of an equation.
@property (copy, readonly) WordMathBox *box;  // Returns a math box object to which you can apply properties.
@property (copy, readonly) WordMathDelimiter *delimiter;  // Returns an math delimiter object that represents the delimiter function
@property (copy, readonly) WordMathEquationArray *equationArray;  // Returns a math equation array object that represents an equation array function.
@property (copy, readonly) WordMathFraction *fraction;  // Returns a math fraction object that represents a fraction.
@property (copy, readonly) WordMathFunc *func;  // Returns a math func object that represents a type of mathematical function.
@property (readonly) WordWdOMathFunctionType functionType;  // Gets the type of the function.
@property (copy, readonly) WordMathGroupChar *groupChar;  // Returns a math group char object that represents a horizontal character placed above or below text in an equation.
@property (copy, readonly) WordMathLeftScripts *leftScripts;
@property (copy, readonly) WordMathLowerLimit *lowerLimit;  // Returns a math lower limit object that represents the lower limit for a function.
@property (copy, readonly) WordMathObject *mathObj;  // Returns an OMath object that represents the equation.
@property (copy, readonly) WordMathMatrix *matrix;  // Returns a math matrix object that represents a mathematical matrix.
@property (copy, readonly) WordMathNary *nary;  // Returns a math nary object that represents the n-ary operation.
@property (copy, readonly) WordMathBar *overbar;  // Returns a math bar object that represents the mathematical overbar for an object.
@property (copy, readonly) WordMathPhantom *phantom;  // Returns a math phantom object that represents an object used for advanced layout of an equation.
@property (copy, readonly) WordMathRadical *radical;  // Returns a math radical object that represents the mathematical radical function.
@property (copy, readonly) WordMathSubAndSuperScript *subAndSuperScript;  // Returns a math sub and super script object that represents a mathematical subscript-superscript object that consists of a base, a subscript, and a superscript.
@property (copy, readonly) WordMathSubscript *subscript;  // Returns a math subscript object that represents the mathematical subscript function.
@property (copy, readonly) WordMathSuperscript *superscript;  // Returns a math superscript object that represents the mathematical superscript function.
@property (copy, readonly) WordTextRange *textRange;  // Returns a text range object that represents the portion of a document that is contained in the math function.
@property (copy, readonly) WordMathUpperLimit *upperLimit;  // Returns a math upper limit object that represents upper limit function.


@end

// Represents a group character object, consisting of a character drawn above or below text, often with the purpose of visually grouping items.
@interface WordMathGroupChar : SBObject <WordGenericMethods>

@property BOOL alignTop;  // Returns or sets whether the grouping character is aligned vertically with the surrounding text or whether the base text that is either above or below the grouping character is aligned vertically with the surrounding text.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isOnTop;  // Gets or sets whether the grouping character is placed above the base text of the group character object. False displays the group character under the base text.
@property NSInteger theCharacter;  // Gets or sets an integer that represents the character placed above or below text in a group character object.


@end

// Represents an equation that contains a superscript or subscript to the left of the base.
@interface WordMathLeftScripts : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *subscript;  // Returns a math object object that represents the subscript for a pre-sub-superscript object.
@property (copy, readonly) WordMathObject *superscript;  // Returns a math object object that represents the superscript for a pre-sub-superscript object.

- (WordMathFunction *) convertLeftScriptsToSubAndSuperScripts;  // Converts an equation with a superscript or subscript to the left of the base of the equation to an equation with a base of a superscript or subscript.

@end

// Represents the lower limit mathematical construct, consisting of text on the baseline and reduced-size text immediately below it.
@interface WordMathLowerLimit : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *limit;  // Returns a math object object that represents the limit of the lower limit object. 

- (WordMathFunction *) lowerLimitToUpperLimit;

@end

// Represents a matrix column.
@interface WordMathMatrixColumn : SBObject <WordGenericMethods>

- (SBElementArray<WordMathObject *> *) mathObjects;

@property NSInteger columnIndex;  // Gets an integer that represents the ordinal position of a matrix column within the collection of matrix columns.
@property WordWdOMathHorizAlignType horizontalAlignment;  // Gets or sets the horizontal alignment for arguments in a matrix column.


@end

// Represents a matrix row.
@interface WordMathMatrixRow : SBObject <WordGenericMethods>

- (SBElementArray<WordMathObject *> *) mathObjects;

@property NSInteger rowIndex;  // Gets an integer that represents the ordinal position of a matrix row within the collection of matrix rows.


@end

// Represents an equation matrix.
@interface WordMathMatrix : SBObject <WordGenericMethods>

- (SBElementArray<WordMathMatrixRow *> *) mathMatrixRows;
- (SBElementArray<WordMathMatrixColumn *> *) mathMatrixColumns;

@property NSInteger columnGap;  // Gets or sets an integer that represents the spacing between columns in a matrix.
@property WordWdOMathSpacingRule columnGapRule;  // Gets or sets the spacing rule for the space that appears between columns in a matrix.
@property NSInteger columnSpacing;  // Gets or sets an integer that represents the spacing for columns in a matrix.
@property BOOL hidePlaceholders;  // Gets or sets whether placeholders in a matrix are hidden from display.
@property NSInteger rowSpacing;  // Gets or sets an integer that represents the spacing for rows in a matrix. 
@property WordWdOMathSpacingRule rowSpacingRule;  // Gets or sets the spacing rule for rows in a matrix.
@property WordWdOMathVertAlignType verticalAlignment;  // Gets or sets the vertical alignment for a matrix.

- (WordMathMatrixColumn *) addMatrixColumnBeforeColumn:(WordMathMatrixColumn *)beforeColumn;  // Creates a matrix column and adds it to a matrix and returns a math matrix column object.
- (WordMathMatrixRow *) addMatrixRowBeforeRow:(WordMathMatrixRow *)beforeRow;  // Creates a matrix row and adds it to a matrix and returns a math matrix row object.

@end

// Represents the mathematical n-ary object, consisting of an n-ary object, a base/operand, and optional upper limits and lower limits.
@interface WordMathNary : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hidesLowerLimit;  // Gets or sets whether to hide the lower limit of an n-ary operator.
@property BOOL hidesUpperLimit;  // Gets or sets whether to hide the upper limit of an n-ary operator.
@property (copy, readonly) WordMathObject *lowerLimit;  // Returns a math object object that represents the lower limit of an n-ary operator.
@property NSInteger operatorCharacter;  // Gets or sets an integer that represents a character used as the n-ary operator.
@property BOOL shouldGrow;  // Gets or sets whether n-ary operators grow to the full height of the arguments that they contain.
@property (copy, readonly) WordMathObject *upperLimit;  // Returns a math object object that represents the upper limit of an n-ary operator.
@property BOOL useSubAndSuperScriptPositioning;  // Gets or sets the positioning of n-ary limits in the subscript-superscript or upper limit-lower limit position.


@end

// Represents an equation
@interface WordMathObject : SBObject <WordGenericMethods>

- (SBElementArray<WordMathFunction *> *) mathFunctions;
- (SBElementArray<WordMathBreak *> *) mathBreaks;

@property NSInteger alignPoint;  // Gets or sets an integer that represents the character position of the alignment point in the equation.
@property (readonly) NSInteger argumentIndex;  // Gets an integer that specifies the argument index of this component relative to the containing math object.
@property (copy, readonly) WordMathFunction *containingFunction;  // Gets the math function that represents the parent, or containing, function.
@property WordWdOMathType displayType;  // Sets or returns the display type of the equation.
@property WordWdOMathJc justification;  // Gets or sets the justification for the equation.
@property (readonly) NSInteger nestingLevel;  // Returns an integer that represents the nesting level for the math object.
@property (copy, readonly) WordMathObject *parentArgument;  // Gets a math object that represents the parent, or containing, argument.
@property (copy, readonly) WordMathMatrixColumn *parentColumn;  // Gets a math matrix column object that represents the parent column in a matrix.
@property (copy, readonly) WordMathObject *parentObject;  // Gets the math object that represents the parent object.
@property (copy, readonly) WordMathMatrixRow *parentRow;  // Gets a math matrix row object that represents the parent row in a matrix.
@property NSInteger scriptSize;  // Gets or sets an integer that represents the script size of an argument, for example, text, script, or script-script.
@property (copy, readonly) WordTextRange *textRange;  // Gets a text range object that represents the portion of a document that contains the equation.

- (void) addMathFunctionInLocation:(WordTextRange *)inLocation mathFunctionType:(WordWdOMathFunctionType)mathFunctionType numberOfArguments:(NSInteger)numberOfArguments numberOfColumns:(NSInteger)numberOfColumns;  // Inserts a new structure, such as a fraction, into an equation at the specified position.
- (void) buildUp;  // Converts an equation to professional/built up format.
- (void) convertToLiteralText;  // Converts an equation to literal text.
- (void) convertToMathText;  // Converts an equation to math text.
- (void) convertToNormalText;  // Converts an equation to normal text.
- (WordMathBreak *) insertMathBreakAtRange:(WordTextRange *)atRange;  // Inserts a break into an equation and returns a math break object that represents the break.
- (void) linearize;  // Converts an equation to linear/built down format.

@end

// Represents a phantom object, which has two primary uses: 1. adding the spacing of the phantom base without displaying that base or 2. suppressing part of the glyph from spacing considerations.
@interface WordMathPhantom : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL showsInvisibles;  // Gets or sets whether the contents of a phantom object are visible.
@property BOOL smash;  // Gets or sets if the contents of the phantom are visible but that the height is not taken into account in the spacing of the layout.
@property BOOL transparent;  // Gets or sets whether a phantom object is transparent.
@property BOOL zeroAscent;  // Gets or sets whether the ascent of the phantom contents is ignored in the spacing of the layout.
@property BOOL zeroDescent;  // Gets or sets whether the descent of the phantom contents is ignored in the spacing of the layout.
@property BOOL zeroWidth;  // Gets or sets whether the width of a phantom object is ignored in the spacing of the layout.


@end

// Represents the mathematical radical object, consisting of a radical, a base, and an optional degree.
@interface WordMathRadical : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *degree;  // Returns a math object object that represents the degree for a radical.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hideDegree;  // Gets or sets whether to hide the degree for the radical.


@end

@interface WordMathRecognizedFunction : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Gets an integer that represents the position of an item in the collection.
@property (copy, readonly) NSString *name;  // Gets a string that represents the name of an equation recognized function.


@end

// Represents an equation with a base that contains a superscript or subscript.
@interface WordMathSubAndSuperScript : SBObject <WordGenericMethods>

@property BOOL alignScripts;  // Gets or sets whether to horizontally align subscripts and superscripts in the sub-superscript object.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *subscript;  // Returns a math object object that represents the subscript for a subscript-superscript object.
@property (copy, readonly) WordMathObject *superscript;  // Returns a math object object that represents the superscript for a subscript-superscript object.

- (WordMathFunction *) convertSubAndSuperScriptsToLeftScripts;  // Converts an equation with a base superscript or subscript to an equation with a superscript or subscript to the left of the base.
- (WordMathFunction *) removeSubscript;  // Removes the subscript for an equation and returns a math function object that represents the updated equation without the subscript.
- (WordMathFunction *) removeSuperscript;  // Removes the superscript for an equation and returns a math function object that represents the updated equation without the superscript.

@end

// Represents an equation with a base that contains a subscript.
@interface WordMathSubscript : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *subscript;  // Returns a math object that represents the subscript for a subscript object.


@end

// Represents an equation with a base that contains a superscript.
@interface WordMathSuperscript : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *superscript;  // Returns a math object object that represents the superscript for a superscript object.


@end

// Represents the upper limit mathematical construct, consisting of text on the baseline and reduced-size text immediately above it.
@interface WordMathUpperLimit : SBObject <WordGenericMethods>

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *limit;  // Returns a math object object that represents the limit of the upper limit object.

- (WordMathFunction *) upperLimitToLowerLimit;

@end

// Represents options associated with page number objects
@interface WordPageNumberOptions : SBObject <WordGenericMethods>

@property WordWdSeparatorType chapterPageSeparator;  // Returns or sets the separator character used between the chapter number and the page number. 
@property NSInteger headingLevelForChapter;  // Returns or sets the heading level style that's applied to the chapter titles in the document. Can be a number from zero through 8, corresponding to heading levels 1 through 9.
@property BOOL includeChapterNumber;  // Returns or sets if a chapter number is included with page numbers.
@property WordWdPageNumberStyle numberStyle;  // Returns or sets the number style for the page number objects.
@property BOOL restartNumberingAtSection;  // Returns or sets if page numbering starts at 1 again at the beginning of the specified section.
@property BOOL showFirstPageNumber;  // Returns or sets if the page number appears on the first page in the section.
@property NSInteger startingNumber;  // Returns or sets the starting page number.


@end

// Represents a page number in a header or footer.
@interface WordPageNumber : SBObject <WordGenericMethods>

@property WordWdPageNumberAlignment alignment;  // Returns or sets a constant that represents the alignment for the page number.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.


@end

// Represents the page setup description.
@interface WordPageSetup : SBObject <WordGenericMethods>

- (SBElementArray<WordTextColumn *> *) textColumns;

@property double bottomMargin;  // Returns or sets the distance in points between the bottom edge of the page and the bottom boundary of the body text.
@property double charsLine;  // Returns or sets the number of characters per line in the document grid.
@property BOOL differentFirstPageHeaderFooter;  // Returns or set if a different header or footer is used on the first page.
@property WordWdPaperTray firstPageTray;  // Returns or sets the paper tray to use for the first page of a document or section.
@property double footerDistance;  // Returns or sets the distance in points between the footer and the bottom of the page.
@property double gutter;  // Returns or sets the amount in points of extra margin space added to each page in a document or section for binding. 
@property WordWdGutterStyle gutterPosition;  // Returns or sets on which side the gutter appears in a document.
@property double headerDistance;  // Returns or sets the distance in points between the header and the top of the page.
@property WordWdLayoutMode layoutMode;  // Returns or sets the layout mode for the current document.
@property double leftMargin;  // Returns or sets the distance in points between the left edge of the page and the left boundary of the body text.
@property BOOL lineBetweenTextColumns;  // Returns or sets if vertical lines appear between all the columns.
@property (copy) WordLineNumbering *lineNumbering;  // Returns or sets the line numbering object that represents the line numbers for the specified page setup object.
@property double linesPage;  // Returns or sets the number of lines per page in the document grid.
@property BOOL mirrorMargins;  // Returns or sets if the inside and outside margins of facing pages are the same width.
@property BOOL oddAndEvenPagesHeaderFooter;  // Returns or sets if the specified page setup object has different headers and footers for odd-numbered and even-numbered pages.
@property WordWdOrientation orientation;  // Returns or sets the orientation of the page.
@property WordWdPaperTray otherPagesTray;  // Returns or sets the paper tray to be used for all but the first page of a document or section.
@property double pageHeight;  // Returns or sets the height of the page in points.
@property double pageWidth;  // Returns or sets the width of the page in points.
@property WordWdPaperSize paperSize;  // Returns or sets the paper size.
@property double rightMargin;  // Returns or sets the distance in points between the right edge of the page and the right boundary of the body text.
@property WordWdSectionStart sectionStart;  // Returns or sets the type of section break for the specified object.
@property BOOL showGrid;  // Determines whether to show the grid.
@property double spacingBetweenTextColumns;  // Returns or sets the spacing in points between columns.
@property BOOL suppressEndnotes;  // Returns or sets if endnotes are printed at the end of the next section that doesn't suppress endnotes. Suppressed endnotes are printed before the endnotes in that section.
@property BOOL textColumnsEvenlySpaced;  // Returns or sets if text columns are evenly spaced.
@property double topMargin;  // Returns or sets the distance in points between the top edge of the page and the top boundary of the body text.
@property WordWdVerticalAlignment verticalAlignment;  // Returns or sets the vertical alignment of text on each page in a document or section.
@property double widthOfTextColumns;  // Returns or sets the width of all text columns

- (void) setAsPageSetupTemplateDefault;  // Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.
- (void) setNumberOfTextColumnsNumberOfColumns:(NSInteger)numberOfColumns;  // Arranges text into the specified number of text columns. 
- (void) togglePortrait;  // Switches between portrait and landscape page orientations for a document or section.

@end

// Represents a window pane.
@interface WordPane : SBObject <WordGenericMethods>

@property BOOL browseToWindow;  // Returns or sets if lines wrap at the right edge of the pane rather than at the right margin of the page.
@property (readonly) NSInteger browseWidth;  // Returns the width in points of the area in which text wraps in the specified pane. This property works only when you're in web layout view.
@property BOOL displayRulers;  // Returns or sets if rulers are displayed for the window
@property BOOL displayVerticalRuler;  // Returns or sets if vertical rulers are displayed for the window
@property (copy, readonly) WordDocument *document;  // Returns a document object associated with this pane.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property NSInteger horizontalPercentScrolled;  // Returns or sets the horizontal scroll position as a percentage of the pane width.
@property NSInteger minimumFontSize;  // Returns or sets the minimum font size in points displayed for the specified pane.
@property (copy, readonly) WordPane *nextPane;  // Returns the next pane object.
@property (copy, readonly) WordPane *previousPane;  // Returns the previous pane object.
@property (copy, readonly) WordSelectionObject *selection;  // Returns the selection object that represents a selected text range or the insertion point.
@property NSInteger verticalPercentScrolled;  // Returns or sets the vertical scroll position as a percentage of the pane width.
@property (copy, readonly) WordView *view;  // Returns a view object that represents the view for the pane.

- (WordZoom *) getZoomZoomType:(WordWdViewType)zoomType;  // Returns a zoom object of the specified type for this pane.

@end

@interface WordRangeEndnoteOptions : SBObject <WordGenericMethods>

@property WordWdEndnoteLocation endnoteLocation;  // Returns or sets the position of endnotes in a range or selection.
@property WordWdNoteNumberStyle endnoteNumberStyle;  // Returns or sets the number style for endnotes in a range or selection.
@property WordWdNumberingRule endnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks.
@property NSInteger endnoteStartingNumber;  // Returns or sets the starting endnote number in a range or selection.


@end

@interface WordRangeFootnoteOptions : SBObject <WordGenericMethods>

@property WordWdFootnoteLocation footnoteLocation;  // Returns or sets the position of footnotes in a range or selection.
@property WordWdNoteNumberStyle footnoteNumberStyle;  // Returns or sets the number style for footnotes in a range or selection.
@property WordWdNumberingRule footnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. 
@property NSInteger footnoteStartingNumber;  // Returns or sets the starting number for footnotes in a range or selection.


@end

// Represents a recently used file.
@interface WordRecentFile : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the recently used file.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the recent file in HFS path style.
@property BOOL readOnly;  // Returns or sets if changes to the document cannot be saved to the original document. 

- (WordDocument *) openRecentFile;  // Opens the recent file and returns a document object.

@end

// Represents the replace criteria for a find-and-replace operation.
@interface WordReplacement : SBObject <WordGenericMethods>

@property (copy) NSString *content;  // Returns or sets the text to replace in the specified range or selection.
@property (copy, readonly) WordFont *fontObject;  // The font object associated with this replacement object.
@property (copy, readonly) WordFrame *frame;  // The frame object associated with this replacement object.
@property NSInteger highlight;  // Returns or sets if highlight formatting is applied to the replacement text.
@property WordWdLanguageID languageID;  // Returns or sets the language for the replacement object
@property WordWdLanguageID languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or set the paragraph format object associated with this replacement object.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style associated with the replacement object.


@end

// Property of View: a person who has made tracked changes in the viewed document.
@interface WordReviewer : SBObject <WordGenericMethods>

@property BOOL visible;


@end

// Represents a change marked with a revision mark.
@interface WordRevision : SBObject <WordGenericMethods>

@property (copy, readonly) NSString *author;  // Returns the name of the user who made the specified tracked change. 
@property (copy, readonly) id cells;
@property (copy, readonly) NSDate *dateValue;  // The date and time that the tracked change was made.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *formatDescription;
@property (copy, readonly) WordTextRange *movedRange;
@property (readonly) WordWdRevisionType revisionType;  // Returns the revision type.
@property (copy, readonly) id style;
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the revision


@end

// Represents the current selection in a window or pane. A selection represents either a selected or highlighted area in the document, or it represents the insertion point if nothing in the document is selected.
@interface WordSelectionObject : SBObject <WordGenericMethods>

- (SBElementArray<WordTable *> *) tables;
- (SBElementArray<WordWord *> *) words;
- (SBElementArray<WordSentence *> *) sentences;
- (SBElementArray<WordCharacter *> *) characters;
- (SBElementArray<WordFootnote *> *) footnotes;
- (SBElementArray<WordEndnote *> *) endnotes;
- (SBElementArray<WordWordComment *> *) WordComments;
- (SBElementArray<WordCell *> *) cells;
- (SBElementArray<WordSection *> *) sections;
- (SBElementArray<WordParagraph *> *) paragraphs;
- (SBElementArray<WordField *> *) fields;
- (SBElementArray<WordFormField *> *) formFields;
- (SBElementArray<WordFrame *> *) frames;
- (SBElementArray<WordBookmark *> *) bookmarks;
- (SBElementArray<WordHyperlinkObject *> *) hyperlinkObjects;
- (SBElementArray<WordColumn *> *) columns;
- (SBElementArray<WordRow *> *) rows;
- (SBElementArray<WordInlineShape *> *) inlineShapes;
- (SBElementArray<WordShape *> *) shapes;
- (SBElementArray<WordMathObject *> *) mathObjects;

@property (readonly) BOOL IPAtEndOfLine;  // Returns true if the insertion point is at the end of a line that wraps to the next line. False if the selection isn't collapsed, if the insertion point isn't at the end of a line, or if the insertion point is positioned before a paragraph mark.
@property (readonly) NSInteger bookmarkId;  // Returns the number of the bookmark that encloses the beginning of the selection. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with the selection
@property (copy, readonly) WordColumnOptions *columnOptions;  // Returns a column options object for this selection.
@property BOOL columnSelectMode;  // Returns or set if column selection mode is active. When this mode is active, the letters COL appear on the status bar.
@property (copy) NSString *content;  // Returns or sets the text in the selection.
@property (copy, readonly) WordDocument *document;  // Returns the document object associated with the selection.
@property (copy, readonly) WordEndnoteOptions *endnoteOptions;  // Returns the endnote options object for the selection.
@property BOOL extendMode;  // Returns or set if extend mode is active.
@property (copy, readonly) WordFind *findObject;  // Returns the find object associated with the selection.
@property double fitTextWidth;  // Returns or sets the width in the current measurement units in which Microsoft Word fits the text in the current selection. 
@property (copy, readonly) WordFont *fontObject;  // The font object associated with the selection.
@property (copy, readonly) WordFootnoteOptions *footnoteOptions;  // Returns the footnote options object for the selection.
@property (copy) WordTextRange *formattedText;  // Returns or sets a text range object that includes the formatted text in the selection.
@property (copy, readonly) WordHeaderFooter *headerFooterObject;  // Returns the header footer object associated with the selection.
@property (readonly) BOOL isEndOfRowMark;  // Returns true if the selection is collapsed and is located at the end-of-row mark in a table.
@property WordWdLanguageID languageID;  // Returns or sets the language for the selection object
@property WordWdLanguageID languageIDEastAsian;  // Returns or sets an east asian language for the selection.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property WordWdTextOrientation orientation;  // Returns or sets the orientation of text in the selection when the text direction feature is enabled.
@property (copy) WordPageSetup *pageSetup;  // Returns or set the page setup object associated with the selection.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or set the paragraph object associated with the selection.
@property (readonly) NSInteger previousBookmarkId;  // Returns the number of the last bookmark that starts before or at the same place as the selection, It returns zero if there's no corresponding bookmark.
@property (copy, readonly) WordRangeEndnoteOptions *rangeEndnoteOptions;
@property (copy, readonly) WordRangeFootnoteOptions *rangeFootnoteOptions;
@property (copy, readonly) WordRowOptions *rowOptions;
@property NSInteger selectionEnd;  // Returns or sets the ending character position of the selection.
@property WordWdSelectionFlags selectionFlags;  // Returns or sets properties of the selection.
@property (readonly) BOOL selectionIsActive;  // Returns if the selection is active.
@property NSInteger selectionStart;  // Returns or sets the starting character position of the selection.
@property (readonly) WordWdSelectionType selectionType;  // Returns the selection type.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the selection
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are included in the elements of the selection
@property BOOL startIsActive;  // Returns or sets if the beginning of the selection is active. If the selection is not collapsed to an insertion point, either the beginning or the end of the selection is active.
@property (readonly) NSInteger storyLength;  // Returns the number of characters in the story that contains the selection
@property (readonly) WordWdStoryType storyType;  // Returns the story type for the selection.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style associated with the selection object.
@property WordWdLanguageID supplementalLanguageID;  // Returns or sets the language for the selection
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the selection

- (double) calculateSelection;  // Calculates a mathematical expression within a selection.
- (void) copyFormat NS_RETURNS_NOT_RETAINED;  // Copies the character formatting of the first character in the selected text. If a paragraph mark is selected, Word copies paragraph formatting in addition to character formatting.
- (void) createTextbox;  // Adds a default-size text box around the selection. If the selection is an insertion point, this method changes the pointer to a cross-hair pointer so that the user can draw a text box.
- (WordTextRange *) endKeyMove:(WordWdUnits)move extend:(WordWdMovementType)extend;  // Moves or extends the selection to the end of the specified unit. This method returns the new range object or missing value if an error occurred.
- (void) escapeKey;  // Cancels a mode such as extend or column select.
- (id) expandBy:(WordWdUnits)by;  // Expands the selection. Returns the number of characters added to the range or selection.
- (void) extendCharacter:(NSString *)character;  // Extends the selection to the next larger unit of text. The progression of selected units of text is as follows: word, sentence, paragraph, section, entire document. The selection is extended by moving the active end of the selection.
- (WordField *) getNextField;  // Returns the next field object.
- (WordField *) getPreviousField;  // Returns the previous field object.
- (NSString *) getSelectionInformationInformationType:(WordWdInformation)informationType;  // Returns information about the selection. 
- (NSInteger) homeKeyMove:(WordWdUnits)move extend:(WordWdMovementType)extend;  // Moves or extends the selection to the beginning of the specified unit. This method returns the new range object or missing value if an error occurred.
- (void) insertCellsShiftCells:(WordWdInsertCells)shiftCells;  // Adds cells to an existing table. The number of cells inserted is equal to the number of cells in the selection.
- (void) insertColumnsPosition:(WordWdInsertRightLeft)position;  // Inserts columns to the left of the column that contains the selection. If the selection isn't in a table, an error occurs. The number of columns inserted is equal to the number of columns selected.
- (void) insertFormulaFormula:(NSString *)formula numberFormat:(NSString *)numberFormat;  // Inserts an = Formula field that contains a formula at the selection. The formula replaces the selection, if the selection isn't collapsed.
- (void) insertRowsPosition:(WordWdInsertAboveBelow)position numberOfRows:(NSInteger)numberOfRows;  // Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.
- (WordRevision *) nextRevisionWrap:(BOOL)wrap;  // Locates and returns the next tracked change as a revision object. The changed text becomes the current selection. Use the properties of the resulting revision object to see what type of change it is, who made it, and so forth.
- (void) pasteFormat;  // Applies formatting copied with the copy format method to the selection. If a paragraph mark was selected when the copy format method was used, Word applies paragraph formatting in addition to character formatting.
- (void) pasteObject;
- (WordRevision *) previousRevisionWrap:(BOOL)wrap;  // Locates and returns the previous tracked change as a revision object. The changed text becomes the current selection. Use the properties of the resulting revision object to see what type of change it is, who made it, and so forth.
- (void) selectCell;  // Selects the entire cell containing the current selection. To use this method, the current selection must be contained within a single cell.
- (void) selectColumn;  // Selects the column that contains the insertion point, or selects all columns that contain the selection. If the selection isn't in a table, an error occurs.
- (void) selectCurrentAlignment;  // Extends the selection forward until text with a different paragraph alignment is encountered.
- (void) selectCurrentColor;  // Extends the selection forward until text with a different color is encountered.
- (void) selectCurrentFont;  // Extends the selection forward until text in a different font or font size is encountered.
- (void) selectCurrentIndent;  // Extends the selection forward until text with different left or right paragraph indents is encountered.
- (void) selectCurrentSpacing;  // Extends the selection forward until a paragraph with different line spacing is encountered.
- (void) selectCurrentTabs;  // Extends the selection forward until a paragraph with different tab stops is encountered.
- (void) selectRow;  // Selects the row that contains the insertion point, or selects all rows that contain the selection. If the selection isn't in a table, an error occurs.
- (void) shrinkDiscontiguousSelection;  // Deselects all but the most recently selected text when a selection contains multiple, unconnected selections.
- (void) shrinkSelection;  // Shrinks the selection to the next smaller unit of text. The progression is as follows: entire document, section, paragraph, sentence, word, insertion point.
- (void) speakText;  // Have the selected text be spoken.
- (void) splitTableInSelection;  // Inserts an empty paragraph above the first row in the selection. If the selection isn't in the first row of the table, the table is split into two tables.
- (void) typeBackspace;  // Deletes the character preceding a collapsed selection, an insertion point. If the selection isn't collapsed to an insertion point, the selection is deleted.
- (void) typeParagraph;  // Inserts a new, blank paragraph. If the selection isn't collapsed to an insertion point, it's replaced by the new paragraph.
- (void) typeTextText:(NSString *)text;  // Inserts the specified text. If the Word options property replace selection is true, the selection is replaced by the specified text. If replace selection is false, the specified text is inserted before the selection.

@end

// Represents a subdocument within a document or range.
@interface WordSubdocument : SBObject <WordGenericMethods>

@property (readonly) BOOL hasFile;  // Returns true if the specified subdocument has been saved to a file.
@property (readonly) NSInteger level;  // Returns the heading level used to create the subdocument.
@property BOOL locked;  // Returns or sets if a subdocument in a master document is locked.
@property (copy, readonly) NSString *name;  // The name of the subdocument.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified subdocument in HFS path style.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the subdocument.

- (WordDocument *) openSubdocument;  // Opens the specified subdocument. Returns a document object representing the opened object.
- (void) splitSubdocumentTextRange:(WordTextRange *)textRange;  // Divides an existing subdocument into two subdocuments at the same level in master document view or outline view. The division is at the beginning of the specified range.

@end

// Contains information about the computer system.
@interface WordSystemObject : SBObject <WordGenericMethods>

@property (readonly) WordWdCountry country;  // Returns the country or region designation of the system
@property WordWdCursorType cursor;  // Returns or sets the state of the pointer.
@property (readonly) NSInteger horizontalResolution;  // Returns the horizontal display resolution, in pixels.
@property (copy, readonly) NSString *operatingSystem;  // Returns the name of the current operating system.
@property (copy, readonly) NSString *processorType;  // Returns the type of processor that the system is using.
@property (copy, readonly) NSString *systemVersion;  // Returns the version number of the operating system.
@property (readonly) NSInteger verticalResolution;  // Returns the vertical display resolution, in pixels.

- (NSString *) getPrivateProfileStringFileName:(NSString *)fileName section:(NSString *)section key:(NSString *)key;  // Returns a string in a settings file.
- (NSString *) getProfileStringSection:(NSString *)section key:(NSString *)key;  // Returns a string from the application's settings file.
- (void) setPrivateProfileStringFileName:(NSString *)fileName section:(NSString *)section key:(NSString *)key privateProfileString:(NSString *)privateProfileString;  // Sets a string in a settings file.
- (void) setProfileStringSection:(NSString *)section key:(NSString *)key profileString:(NSString *)profileString;  // Sets a string from the application's settings file.

@end

// Represents a single tab stop. 
@interface WordTabStop : SBObject <WordGenericMethods>

@property WordWdTabAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified tab stop.
@property (readonly) BOOL customTab;  // Returns if the specified tab stop is a custom tab stop.
@property (copy, readonly) WordTabStop *nextTabStop;  // Returns the next tab stop object
@property (copy, readonly) WordTabStop *previousTabStop;  // Returns the previous tab stop object
@property WordWdTabLeader tabLeader;  // Returns or sets the leader for the specified tab stop object
@property double tabStopPosition;  // Returns or sets the position of a tab stop relative to the left margin.


@end

// Represents a single table of authorities in a document
@interface WordTableOfAuthorities : SBObject <WordGenericMethods>

@property NSInteger category;  // Returns or sets the category of entries to be included in a table of authorities.
@property (copy) NSString *entrySeparator;  // Returns or sets the characters up to five that separate a table of authorities entry and its page number. The default is a tab character with a dotted leader.
@property BOOL includeCategoryHeader;  // Returns or sets if the category name for a group of entries appears in the table of authorities.
@property (copy) NSString *includeSequenceName;  // Returns or sets the Sequence SEQ field identifier for a table of authorities.
@property BOOL keepEntryFormatting;  // Returns or sets if formatting from table of authorities entries is applied to the entries in the specified table of authorities.
@property (copy) NSString *pageNumberSeparator;  // Returns of sets the characters up to five that separate individual page references in a table of authorities. The default is a comma and a space.
@property (copy) NSString *pageRangeSeparator;  // Returns or sets the characters up to five that separate a range of pages in a table of authorities. The default is an en dash.
@property BOOL passim;  // Returns or sets if five or more page references to the same authority are replaced with Passim.
@property (copy) NSString *separator;  // Returns or sets the characters up to five between the sequence number and the page number. A hyphen is the default character.
@property WordWdTabLeader tabLeader;  // Returns or sets the character between entries and their page numbers in a table of authorities.
@property (copy) NSString *tableOfAuthoritiesBookmark;  // Returns or sets the name of the bookmark from which to collect table of authorities entries.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the table of authorities object.


@end

// Represents a single table of contents in a document.
@interface WordTableOfContents : SBObject <WordGenericMethods>

- (SBElementArray<WordHeadingStyle *> *) headingStyles;

@property BOOL includePageNumbers;  // Returns or sets if page numbers are included in the table of contents.
@property NSInteger lowerHeadingLevel;  // Returns or sets the ending heading level for a table of contents.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in a table of contents.
@property WordWdTabLeader tabLeader;  // Returns or sets the character between entries and their page numbers in a table of contents.
@property (copy) NSString *tableId;  // Returns or sets a one-letter identifier that's used to build a table of contents or table of figures from TOC fields.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the table of contents object.
@property NSInteger upperHeadingLevel;  // Returns or sets the starting heading level for a table of contents.
@property BOOL useFields;  // Returns or sets if table of contents entry fields are used to create a table of contents.
@property BOOL useHeadingStyles;  // Returns or sets if built-in heading styles are used to create a table of contents.


@end

// Represents a single table of figures in a document.
@interface WordTableOfFigures : SBObject <WordGenericMethods>

- (SBElementArray<WordHeadingStyle *> *) headingStyles;

@property (copy) NSString *caption;  // Returns or sets the label that identifies the items to be included in a table of figures.
@property BOOL includeLabel;  // Returns or sets if the caption label and caption number are included in a table of figures.
@property BOOL includePageNumbers;  // Returns or sets if page numbers are included in the table of figures.
@property NSInteger lowerHeadingLevel;  // Returns or sets the ending heading level for a table of figures.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in a table of figures.
@property WordWdTabLeader tabLeader;  // Returns or sets the character between entries and their page numbers in a table of figures.
@property (copy) NSString *tableId;  // Returns or sets a one-letter identifier that's used to build a table of figures from TOC fields.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the table of figures object.
@property NSInteger upperHeadingLevel;  // Returns or sets the starting heading level for a table of figures.
@property BOOL useFields;  // Returns or sets if table of figures entry fields are used to create a table of figures.
@property BOOL useHeadingStyles;  // Returns or sets if built-in heading styles are used to create a table of figures.


@end

// Represents a document template.
@interface WordTemplate : SBObject <WordGenericMethods>

- (SBElementArray<WordAutoTextEntry *> *) autoTextEntries;
- (SBElementArray<WordDocumentProperty *> *) documentProperties;
- (SBElementArray<WordCustomDocumentProperty *> *) customDocumentProperties;
- (SBElementArray<WordListTemplate *> *) listTemplates;
- (SBElementArray<WordBuildingBlock *> *) buildingBlocks;
- (SBElementArray<WordBuildingBlockType *> *) buildingBlockTypes;

@property WordWdFarEastLineBreakLevel eastAsianLineBreak;  // Returns or sets the line break control level for the specified template. This property is ignored if the Ôøíeast asian line break controlÔøì property is set to false. Note that Ôøíeast asian line break controlÔøì is a paragraph format property.
@property (copy, readonly) NSString *fullName;  // Specifies the name of the template including the drive or Web path in HFS path style.
@property WordWdJustificationMode justificationMode;  // Returns or sets the character spacing adjustment for the specified template.
@property WordWdLanguageID languageID;  // Returns or sets the language for the template object
@property WordWdLanguageID languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (copy, readonly) NSString *name;  // Returns the name of the template.
@property (copy) NSString *noLineBreakAfter;  // Returns or sets the kinsoku characters after which Microsoft Word will not break a line.
@property (copy) NSString *noLineBreakBefore;  // Returns or sets the kinsoku characters before which Microsoft Word will not break a line.
@property BOOL noProofing;  // Returns or sets if the spelling and grammar checker should ignore documents based on this template.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the template file in HFS path style.
@property (copy, readonly) NSString *posixFullName;  // Specifies the name of the template including the drive or Web path in POSIX path style.
@property BOOL saved;  // Return or set if the template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed.
@property (readonly) WordWdTemplateType templateType;  // Returns the template type.

- (WordBuildingBlock *) addBuildingBlockToTemplateName:(NSString *)name type:(WordWdBuildingBlockTypes)type category:(NSString *)category fromLocation:(WordTextRange *)fromLocation objectDescription:(NSString *)objectDescription insertOptions:(WordWdDocPartInsertOptions)insertOptions;  // Creates a new building block entry in a template and returns a building block object that represents the new building block entry.
- (WordAutoTextEntry *) appendToSpikeRange:(WordTextRange *)range;  // Deletes the specified range and adds the contents of the range to the Spike which is a built-in autotext entry.
- (WordDocument *) openAsDocument;  // Opens the specified template as a document and returns a document object. Opening a template as a document allows the user to edit the contents of the template.

@end

// Represents a single text column.
@interface WordTextColumn : SBObject <WordGenericMethods>

@property double spaceAfter;  // Returns or sets the amount of spacing in points after the text column.
@property double width;  // Returns or sets the width of the object.


@end

// Represents a single text form field.
@interface WordTextInput : SBObject <WordGenericMethods>

@property (copy) NSString *defaultTextInput;  // Returns or sets the text that represents the default text box contents.
@property (copy, readonly) NSString *format;  // Returns the formatting string for this text input object.
@property (readonly) WordWdTextFormFieldType textInputFieldType;  // Returns the type of text form field.
@property (readonly) BOOL valid;  // Returns if the text input object is valid.
@property NSInteger width;  // Returns or sets the width of the object.

- (void) editTypeFormFieldType:(WordWdTextFormFieldType)formFieldType defaultType:(NSString *)defaultType typeFormat:(NSString *)typeFormat enabled:(BOOL)enabled;  // Sets options for the specified text form field.

@end

// Represents options that control how text is retrieved from a text range object.
@interface WordTextRetrievalMode : SBObject <WordGenericMethods>

@property BOOL includeFieldCodes;  // Returns or sets if the text retrieved from the specified range includes field codes.
@property BOOL includeHiddenText;  // Returns or sets if the text retrieved from the specified range includes hidden text.
@property WordWdViewType viewType;  // Returns or sets the view for the text retrieval mode object.


@end

// Represents a variable stored as part of a document. Document variables are used to preserve macro settings in between macro sessions.
@interface WordVariable : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the variable
@property (copy) NSString *variableValue;  // Returns or sets the value of a variable as a string.


@end

// Contains the view attributes, show all, field shading, table gridlines, and so on, for a window or pane.
@interface WordView : SBObject <WordGenericMethods>

- (SBElementArray<WordReviewer *> *) reviewers;

@property BOOL browseToWindow;  // Returns or sets if lines wrap at the right edge of the window rather than at the right margin of the page.
@property BOOL conflictMode;  // Returns or sets the Conflict Mode.
@property BOOL dataMergeDataView;  // Returns or sets if data merge data is displayed instead of data merge fields in the specified window.
@property BOOL draft;  // Returns or sets if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.
@property NSInteger enlargeFontsLessThan;  // Returns or sets the point size below which screen fonts are automatically scaled to the larger size.
@property WordWdFieldShading fieldShading;  // Returns or sets on-screen shading for form fields.
@property BOOL fullScreen;  // Returns or sets if the window is in full-screen view.
@property BOOL magnifier;  // Returns or sets if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages.
@property WordWdRevisionsMode revisionsMode;  // Returns or sets the way Microsoft Word shows revisions.
@property WordWdRevisionsView revisionsView;  // Returns or sets whether Microsoft Word shows revisions based on the final or the original document.
@property WordWdSeekView seekView;  // Returns or sets the document element displayed in print layout view.
@property BOOL showAll;  // Returns or sets if all nonprinting characters such as hidden text, tab marks, space marks, and paragraph marks are displayed. 
@property BOOL showBookmarks;  // Returns or sets if square brackets are displayed at the beginning and end of each bookmark.
@property BOOL showComments;  // Returns or sets if Microsoft Word displays comments.
@property BOOL showDrawings;  // Returns or sets if objects created with the drawing tools are displayed in print layout view.
@property BOOL showFieldCodes;  // Returns or sets if field codes are displayed.
@property BOOL showFirstLineOnly;  // Returns or sets if only the first line of body text is shown in outline view.
@property BOOL showFormat;  // Returns or set if character formatting is visible in outline view.
@property BOOL showFormatChanges;  // Returns or sets if Microsoft Word displays formatting changes.
@property BOOL showHiddenText;  // Returns or set if text formatted as hidden text is displayed. 
@property BOOL showHighlight;  // Returns or sets if highlight formatting is displayed and printed with a document.
@property BOOL showHyphens;  // Returns or sets if optional hyphens are displayed. An optional hyphen indicates where to break a word when it falls at the end of a line.
@property BOOL showInsertionsAndDeletions;  // Returns or sets if Microsoft Word displays insertions and deletions.
@property BOOL showMainTextLayer;  // Returns or sets if the text in the specified document is visible when the header and footer areas are displayed.
@property BOOL showObjectAnchors;  // Returns or set if object anchors are displayed next to items that can be positioned in print layout view.
@property BOOL showOptionalBreaks;  // Returns or sets if Microsoft Word displays optional line breaks.
@property BOOL showOtherAuthors;  // Returns or sets whether Microsoft Word shows other authors who are also editing the document.
@property BOOL showParagraphs;  // Returns or sets if paragraph marks are displayed. 
@property BOOL showPicturePlaceHolders;  // Returns or sets if blank boxes are displayed as placeholders for pictures.
@property BOOL showRevisionsAndComments;  // Returns or sets if Microsoft Word displays revisions and comments.
@property BOOL showSpaces;  // Returns or sets if space characters are displayed.
@property BOOL showTabs;  // Returns or sets if tab characters are displayed.
@property BOOL showTextBoundaries;  // Returns or sets if dotted lines are displayed around page margins, text columns, objects, and frames in print layout view. 
@property WordWdSpecialPane splitSpecial;  // Returns or sets the active window pane.
@property BOOL tableGridlines;  // Returns or sets if table gridlines are displayed. 
@property WordWdViewType viewType;  // Returns or sets the view type.
@property BOOL wrapToWindow;  // Returns or sets if lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.
@property (copy, readonly) WordZoom *zoom;  // Returns the zoom object associated with this view object.

- (void) collapseOutlineTextRange:(WordTextRange *)textRange;  // Collapses the text under the selection or the specified range by one heading level.
- (void) expandOutlineTextRange:(WordTextRange *)textRange;  // Expands the text under the selection or the specified range by one heading level.
- (void) nextHeaderFooter;  // If the selection is in a header, this method moves to the next header within the current section or to the first header in the following section. If the selection is in a footer, this method moves to the next footer.
- (void) previousHeaderFooter;  // If the selection is in a header, this method moves to the previous header within the current section or to the last header in the previous section. If the selection is in a footer, this method moves to the previous footer.
- (void) showAllHeadings;  // Toggles between showing all text, headings and body text, and showing only headings.
- (void) showHeadingLevel:(NSInteger)level;  // Shows all headings up to the specified heading level and hides subordinate headings and body text. This method generates an error if the view isn't outline view or master document view.

@end

// Contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page.
@interface WordWebOptions : SBObject <WordGenericMethods>

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page.
@property WordMsoEncoding encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property WordMsoScreenSize screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL useLongFileNames;  // Returns or sets if long file names are used when you save the document as a Web page.

- (void) useDefaultFolderSuffix;  // Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

@end

// Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.
@interface WordWindow (MicrosoftWordSuite)

- (SBElementArray<WordPane *> *) panes;

@property WordWdIMEMode IMEMode;  // Returns or sets the default start-up mode for the Japanese Input Method Editor, IME
@property (readonly) BOOL active;  // Returns if the window is active.
@property (copy, readonly) WordPane *activePane;  // Returns a pane object that represents the active pane for the window 
@property (copy) NSString *caption;  // Returns or sets the caption text for the document window.
@property BOOL displayHorizontalScrollBar;  // Returns or sets if the horizontal scroll bar is visible
@property BOOL displayRulers;  // Returns or sets if rulers are displayed for the window
@property BOOL displayScreenTips;  // Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted.
@property BOOL displayVerticalRuler;  // Returns or sets if vertical rulers are displayed for the window
@property BOOL displayVerticalScrollBar;  // Returns or sets if the vertical scroll bar is visible
@property (copy, readonly) WordDocument *document;  // Returns a document object associated with the pane. 
@property BOOL documentMap;  // Returns or sets if the document map is visible.
@property NSInteger documentMapPercentWidth;  // Returns or sets the width of the document map as a percentage of the width of the specified window.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property NSInteger height;  // Returns or sets the height of the object.
@property NSInteger horizontalPercentScrolled;  // Returns or sets the horizontal scroll position as a percentage of the document width.
@property NSInteger leftPosition;  // Returns or sets the horizontal position of the object.
@property (copy, readonly) WordWindow *nextWindow;  // Returns the next window object
@property (copy, readonly) WordWindow *previousWindow;  // Returns the previous window object
@property (copy, readonly) WordSelectionObject *selection;  // Returns the selection object that represents a selected text range or the insertion point.
@property NSInteger splitVertical;  // Returns or sets the vertical split percentage for the specified window. To remove the split, set this property to zero or set the split window property to false.
@property BOOL splitWindow;  // Returns or sets if the window is split into multiple panes. When setting this to true the window is split into two equal-sized window panes.
@property double styleAreaWidth;  // Returns or sets the width of the style area in points.
@property NSInteger top;  // Returns or sets the vertical position of the object.
@property NSInteger verticalPercentScrolled;  // Returns or sets the vertical scroll position as a percentage of the document width.
@property (copy, readonly) WordView *view;  // Returns a view object that represents the view for the window.
@property NSInteger width;  // Returns or sets the width of the object.
@property (readonly) NSInteger windowNumber;  // Returns the window number of the document displayed in the specified window. For example, if the caption of the window is Sales.doc:2, this property returns the number 2.
@property WordWdWindowState windowState;  // Returns or sets the state of the window.
@property (readonly) WordWdWindowType windowType;  // Returns the window type.

@end

// Contains magnification options, for example, the zoom percentage for a window or pane.
@interface WordZoom : SBObject <WordGenericMethods>

@property NSInteger pageColumns;  // Returns or sets the number of pages to be displayed side by side on-screen at the same time in print layout view or print preview.
@property WordWdPageFit pageFit;  // Returns or sets the view magnification of a window so that either the entire page is visible or the entire width of the page is visible.
@property NSInteger pageRows;  // Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. 
@property NSInteger percentage;  // Returns or sets the magnification for a window as a percentage.


@end



/*
 * Drawing Suite
 */

@interface WordAdjustment : SBObject <WordGenericMethods>

@property double adjustment_value;  // Returns or sets a floating point adjustment for a shape.


@end

// Contains properties and methods that apply to line callouts.
@interface WordCalloutFormat : SBObject <WordGenericMethods>

@property BOOL accent;  // Returns or sets if a vertical accent bar separates the callout text from the callout line.
@property WordMsoCalloutAngleType angle;  // Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box.
@property BOOL autoAttach;  // Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box.
@property (readonly) BOOL autoLength;  // Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false.
@property (readonly) double calloutFormatLength;  // When the AutoLength property of the specified callout is set to False, the Length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box.
@property BOOL calloutHasBorder;  // Returns or sets whether the text in the specified callout is surrounded by a border.
@property WordMsoCalloutType calloutType;  // Returns or sets the callout type.
@property (readonly) double drop;  // For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
@property (readonly) WordMsoCalloutDropType dropType;  // Returns a value that indicates where the callout line attaches to the callout text box.
@property double gap;  // Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box.


@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface WordFillFormat : SBObject <WordGenericMethods>

@property (copy) NSArray<NSNumber *> *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property WordMsoThemeColorIndex backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) WordMsoFillType fillType;  // Returns the shape fill format type.
@property (copy) NSArray<NSNumber *> *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property WordMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) WordMsoGradientColorType gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend.
@property (readonly) WordMsoGradientStyle gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2.
@property (readonly) WordMsoPatternType pattern;  // Returns the value that represents the pattern applied to the specified fill or line.
@property (readonly) WordMsoPresetGradientType presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) WordMsoPresetTexture presetTexture;  // Returns the preset texture for the specified fill.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property (readonly) WordMsoTextureType textureType;  // Returns the texture type for the specified fill.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Represents the glow formatting for a shape or range of shapes.
@interface WordGlowFormat : SBObject <WordGenericMethods>

@property (copy) NSArray<NSNumber *> *color;  // Returns or sets the color for the specified glow format.
@property WordMsoThemeColorIndex colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents horizontal line formatting.
@interface WordHorizontalLineFormat : SBObject <WordGenericMethods>

@property WordWdHorizontalLineAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified horizontal line.
@property BOOL noShade;  // Returns or sets if Microsoft Word draws the specified horizontal line without 3-D shading.
@property double percentWidth;  // Returns or sets the length of the specified horizontal line expressed as a percentage of the window width.
@property WordWdHorizontalLineWidthType widthType;  // Returns or sets the width type for the specified horizontal line format object.


@end

// Represents a graphic object in the text layer of a document.
@interface WordInlineShape : SBObject <WordGenericMethods>

@property (copy) NSString *alternativeText;  // Returns or sets the alternative text associated with a shape in a Web page.
@property (readonly) NSInteger anchorID;  // Returns the anchor id for the specified shape.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with the inline shape
@property (readonly) NSInteger editID;  // Returns the edit id for the specified shape.
@property (copy, readonly) WordField *field;  // Returns a field object that represents the field associated with the specified inline shape.
@property (copy, readonly) WordFillFormat *fillFormat;  // Returns the fill format object associated with this inline shape object.
@property (copy, readonly) WordGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property double height;  // Returns or sets the height of the object.
@property (copy, readonly) WordHorizontalLineFormat *horizontalLineFormat;  // Returns the horizontal line format object associated with this inline shape object.
@property (copy, readonly) WordHyperlinkObject *hyperlink;  // Returns the hyperlink object associated with this inline shape object.
@property double inlineShapeScaleHeight;  // Returns or sets the height of the specified inline shape relative to its original size. 
@property double inlineShapeScaleWidth;  // Returns or sets the width of the specified inline shape relative to its original size. 
@property (readonly) WordWdInlineShapeType inlineShapeType;  // The type of this inline shape.
@property (readonly) BOOL isPictureBullet;  // Returns true if an InlineShape object is a picture bullet.
@property (copy, readonly) WordLineFormat *lineFormat;  // Returns the line format object associated with this inline shape object.
@property (copy, readonly) WordLinkFormat *linkFormat;  // Returns the link format object associated with this inline shape object.
@property BOOL lockAspectRatio;  // Returns or set if the specified shape retains its original proportions when you resize it.
@property (copy, readonly) WordPictureFormat *pictureFormat;  // Returns the picture format object associated with this inline shape object.
@property (copy, readonly) WordReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property (copy, readonly) WordShadowFormat *shadow;  // Returns the shadow format object associated with this shape object.
@property (copy, readonly) WordSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the inline shape.
@property double width;  // Returns or sets the width of the object.
@property (copy, readonly) WordWordArtFormat *wordArtFormat;  // Returns the word art format object associated with the word art shape object.

- (WordShape *) convertToShape;  // Converts an inline shape to a free-floating shape.
- (WordBorder *) getBorderWhichBorder:(WordWdBorderType)whichBorder;  // Returns the specified border object.
- (void) reset;  // Removes changes that were made to an inline shape.

@end

// Represents a horizontal line object in the text layer of a document.
@interface WordInlineHorizontalLine : WordInlineShape

@property (copy, readonly) NSString *fileName;  // The file name that contains the picture of the horizontal line.


@end

// Represents a picture bullet object in the text layer of a document.
@interface WordInlinePictureBullet : WordInlineShape

@property (copy, readonly) NSString *fileName;  // The file name that contains the picture of the picture bullet.


@end

// Represents a picture object in the text layer of a document.
@interface WordInlinePicture : WordInlineShape

@property (copy, readonly) NSString *fileName;  // The file name that contains the picture.
@property (readonly) BOOL linkToFile;  // Returns true if the picture shape is linked to the file.
@property (copy, readonly) WordPictureFormat *pictureFormat;  // Returns the picture format object associated with this inline picture object.
@property (readonly) BOOL saveWithDocument;  // Specifies if the picture should be saved with the document.


@end

// Represents line and arrowhead formatting. For a line, the line format object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.
@interface WordLineFormat : SBObject <WordGenericMethods>

@property (copy) NSArray<NSNumber *> *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property WordMsoThemeColorIndex backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property WordMsoArrowheadLength beginArrowheadLength;  // Returns or sets the length of the arrowhead at the beginning of the specified line.
@property WordMsoArrowheadStyle beginArrowheadStyle;  // Returns or sets the style of the arrowhead at the beginning of the specified line.
@property WordMsoArrowheadWidth beginArrowheadWidth;  // Returns or sets the width of the arrowhead at the beginning of the specified line.
@property WordMsoLineDashStyle dashStyle;  // Returns or sets the dash style for the specified line.
@property WordMsoArrowheadLength endArrowheadLength;  // Returns or sets the length of the arrowhead at the end of the specified line.
@property WordMsoArrowheadStyle endArrowheadStyle;  // Returns or sets the style of the arrowhead at the end of the specified line.
@property WordMsoArrowheadWidth endArrowheadWidth;  // Returns or sets the width of the arrowhead at the end of the specified line.
@property (copy) NSArray<NSNumber *> *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property WordMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property WordMsoLineStyle lineStyle;  // Returns or sets the line format style.
@property WordMsoPatternType pattern;  // Returns or sets a value that represents the pattern applied to the specified fill or line.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.
@property double weight;  // Returns or sets the thickness of the specified line in points.


@end

// Represents a Microsoft Office theme.
@interface WordOfficeTheme : SBObject <WordGenericMethods>

@property (copy, readonly) WordThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) WordThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) WordThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

// Contains properties and methods that apply to pictures. 
@interface WordPictureFormat : SBObject <WordGenericMethods>

@property double brightness;  // Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) id transparencyColor;  // Returns or sets the transparent color for the specified picture as a RGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

// Represents the reflection effect in Office graphics.
@interface WordReflectionFormat : SBObject <WordGenericMethods>

@property WordMsoReflectionType reflectionType;  // Returns or sets the type of the reflection format object.


@end

// Represents shadow formatting for a shape.
@interface WordShadowFormat : SBObject <WordGenericMethods>

@property double blur;  // Returns or sets the blur, in points, of the specified shadow.
@property (copy) NSArray<NSNumber *> *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property WordMsoThemeColorIndex foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;  // Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
@property double offsetX;  // Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
@property double offsetY;  // Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape.
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property WordMsoShadowStyle shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property WordMsoShadowType shadowType;  // Returns or sets the shape shadow type.
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Represents an object in the drawing layer.
@interface WordShape : SBObject <WordGenericMethods>

- (SBElementArray<WordAdjustment *> *) adjustments;

@property (copy, readonly) WordTextRange *anchor;  // Returns a text range object that represents the anchoring range for the specified shape.
@property (readonly) NSInteger anchorID;  // Returns the anchor id for the specified shape.
@property WordMsoAutoShapeType autoShapeType;  // Returns or sets the shape type for the specified shape object, which must represent an autoshape.
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger editID;  // Returns the edit id for the specified shape.
@property (copy, readonly) WordFillFormat *fillFormat;  // Return the fill format object associated with this shape object.
@property (copy, readonly) WordGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property BOOL hasChart;  // True if the specified shape has a chart.
@property double height;  // Returns or sets the height of the object.
@property (readonly) BOOL horizontalFlip;  // Returns true if the shape has been flipped horizontally. 
@property (copy, readonly) WordHyperlinkObject *hyperlink;  // Returns the hyperlink object associated with this shape object.
@property double leftPosition;  // Returns or sets the horizontal position of the object.
@property (copy, readonly) WordLineFormat *lineFormat;  // Returns the line format object associated with this shape object.
@property (copy, readonly) WordLinkFormat *linkFormat;  // Returns the link format object associated with this shape object.
@property BOOL lockAnchor;  // Returns or sets if the specified shape object's anchor is locked to the anchoring range. When a shape has a locked anchor, you cannot move the shape's anchor by dragging it, the anchor doesn't move as the shape is moved.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it.
@property (copy) NSString *name;  // Returns or sets the name of this shape object.
@property (copy, readonly) WordReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property WordWdRelativeHorizontalPosition relativeHorizontalPosition;  // Returns or sets if the horizontal position of the shape is relative.
@property WordWdRelativeVerticalPosition relativeVerticalPosition;  // Returns or sets if the vertical position of the shape is relative.
@property double rotation;  // Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation.
@property (copy, readonly) WordShadowFormat *shadow;  // Returns the shadow format object associated with this shape object.
@property (readonly) WordMsoShapeType shapeType;  // Returns the shape type.
@property (copy, readonly) WordSoftEdgeFormat *softEdgeFormat;  // Returns the formatting properties for a soft edge effect.
@property (copy, readonly) WordTextFrame *textFrame;  // Returns the text frame object associated with this shape object.
@property (copy, readonly) WordThreeDFormat *threeDFormat;  // Returns the threeD format object associated with this shape object.
@property double top;  // Returns or sets the vertical position of the object.
@property (readonly) BOOL verticalFlip;  // Returns true if the shape has been flipped vertically.
@property BOOL visible;  // Returns or sets if the shape object is visible.
@property double width;  // Returns or sets the width of the object.
@property (copy, readonly) WordWrapFormat *wrapFormat;  // Returns the wrap format object associated with this shape object.
@property (readonly) NSInteger zOrderPosition;  // Returns the position of the specified shape in the z-order.

- (void) apply;  // Applies to the specified shape formatting that has been copied using the pick up method.
- (WordFrame *) convertToFrame;  // Converts the specified shape to a frame. Returns a frame object that represents the new frame.
- (WordInlineShape *) convertToInlineShape;  // Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only picture shapes.
- (void) flipFlipCommand:(WordMsoFlipCmd)flipCommand;  // Flips a shape horizontally or vertically.
- (void) pickUp;  // Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.
- (void) rerouteConnections;  // Reroutes the connections between shapes.
- (void) setShapesDefaultProperties;  // Applies the formatting of the specified shape to a default shape for that document. New shapes inherit many of their attributes from the default shape.
- (void) zOrderZOrderCommand:(WordMsoZOrderCmd)zOrderCommand;  // Moves the specified shape in front of or behind other shapes.

@end

// Represents an callout object in the drawing layer.
@interface WordCallout : WordShape

@property (copy, readonly) WordCalloutFormat *calloutFormat;  // Returns the callout format object associated with this callout shape object.
@property (readonly) WordMsoCalloutType calloutType;  // Return the type of this callout


@end

// The line shape uses begin line X, begin line Y, end line X, and end line Y when created
@interface WordLineShape : WordShape

@property double beginLineX;  // Returns or sets the starting X coordinate for the line shape.
@property double beginLineY;  // Returns or sets the starting Y coordinate for the line shape.
@property double endLineX;  // Returns or sets the ending X coordinate for the line shape.
@property double endLineY;  // Returns or sets the ending Y coordinate for the line shape.


@end

// Represents an picture object in the drawing layer.
@interface WordPicture : WordShape

@property (copy, readonly) NSString *fileName;  // The name of the file containing the picture.
@property (readonly) BOOL linkToFile;  // Returns true if the picture shape is linked to the file.
@property (copy, readonly) WordPictureFormat *pictureFormat;  // Returns the picture format object associated with this picture shape.
@property (readonly) BOOL saveWithDocument;  // Specifies if the picture should be saved with the document.

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(WordMsoScaleFrom)scale;  // Scales the height of the picture shape by a specified factor.
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(WordMsoScaleFrom)scale;  // Scales the width of the shape by a specified factor.

@end

// Represents the soft edge formatting for a shape or range of shapes.
@interface WordSoftEdgeFormat : SBObject <WordGenericMethods>

@property WordMsoSoftEdgeType softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Represents a standard horizontal line object in the text layer of a document.
@interface WordStandardInlineHorizontalLine : WordInlineShape


@end

// Represents an text box object in the drawing layer.
@interface WordTextBox : WordShape

@property (readonly) WordMsoTextOrientation textOrientation;  // Returns the orientation of the text inside the text shape.


@end

// Represents the text frame in a shape object. Contains the text in the text frame as well as the properties that control the margins and orientation of the text frame.
@interface WordTextFrame : SBObject <WordGenericMethods>

@property (copy, readonly) WordTextRange *containingRange;  // Returns a text range object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to.
@property (readonly) BOOL hasText;  // Returns true if the specified shape has text associated with it.
@property double marginBottom;  // Returns or sets the distance in points between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance in points between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance in points between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance in points between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property (copy) WordTextFrame *nextTextframe;  // Returns or sets the next text frame object.
@property WordMsoTextOrientation orientation;  // Returns or sets the orientation of the text inside the frame.
@property (readonly) BOOL overflowing;  // Returns if the text inside the specified text frame doesn't all fit within the frame.
@property (copy) WordTextFrame *previousTextframe;  // Returns or sets the previous text frame object.
@property (copy, readonly) WordTextRange *textRange;  // Returns a text range object that represents the text range for this text frame.

- (void) breakForwardLink;  // Breaks the forward link for the specified text frame, if such a link exists.
- (BOOL) validLinkTargetTargetTextframe:(WordTextFrame *)targetTextframe;  // Determines whether the text frame of one shape can be linked to the text frame of another shape. Returns false if target textframe already contains text or is already linked, or if the shape doesn't support attached text.

@end

// Represents the color scheme of a Microsoft Office 2007 theme.
@interface WordThemeColorScheme : SBObject <WordGenericMethods>

- (SBElementArray<WordThemeColor *> *) themeColors;

- (NSArray<NSNumber *> *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface WordThemeColor : SBObject <WordGenericMethods>

@property (copy) NSArray<NSNumber *> *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) WordMsoThemeColorSchemeIndex themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface WordThemeEffectScheme : SBObject <WordGenericMethods>

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file

@end

// Represents the font scheme of a Microsoft Office theme.
@interface WordThemeFontScheme : SBObject <WordGenericMethods>

- (SBElementArray<WordMinorThemeFont *> *) minorThemeFonts;
- (SBElementArray<WordMajorThemeFont *> *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface WordThemeFont : SBObject <WordGenericMethods>

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface WordMajorThemeFont : WordThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface WordMinorThemeFont : WordThemeFont


@end

// Represents a collection of major and minor fonts in the font scheme of a Microsoft Office 2007 theme.
@interface WordThemeFonts : SBObject <WordGenericMethods>


@end

// Represents a shape's three-dimensional formatting.
@interface WordThreeDFormat : SBObject <WordGenericMethods>

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property WordMsoBevelType bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property WordMsoBevelType bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy) NSArray<NSNumber *> *contourColor;  // Returns or sets the color of the contour of an object or text.
@property WordMsoThemeColorIndex contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy) NSArray<NSNumber *> *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property WordMsoThemeColorIndex extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property WordMsoExtrusionColorType extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property WordMsoPresetCamera presetCamera;  // Returns a constant that represents the camera preset.
@property WordMsoPresetExtrusionDirection presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property WordMsoPresetLightingDirection presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property WordMsoLightRigType presetLightingRig;  // Returns a constant that represents the lighting preset.
@property WordMsoPresetLightingSoftness presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property WordMsoPresetMaterial presetMaterial;  // Returns or sets the extrusion surface material.
@property WordMsoPresetThreeDFormat presetThreeDFormat;  // Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Contains properties and methods that apply to WordArt objects.
@interface WordWordArtFormat : SBObject <WordGenericMethods>

@property WordMsoTextEffectAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified text effect.
@property BOOL bold;  // Returns or sets if the text of the word art shape is formatted as bold.
@property (copy) NSString *fontName;  // Returns or sets the font name of the font used by this word art shape.
@property double fontSize;  // Returns or sets the font size of the font used by this word art shape.
@property BOOL italic;  // Returns or sets if the text of the word art shape is formatted as italic.
@property BOOL kernedPairs;  // Returns or sets if character pairs in a WordArt object have been kerned. 
@property BOOL normalizedHeight;  // Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height.
@property WordMsoPresetTextEffectShape presetShape;  // Returns or sets the shape of the specified WordArt.
@property WordMsoPresetTextEffect presetWordArtEffect;  // Returns or sets the style of the specified WordArt.
@property BOOL rotatedChars;  // Returns or sets if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. If false, characters in the specified WordArt retain their original orientation relative to the bounding shape.
@property double tracking;  // Returns or sets the ratio of the horizontal space allotted to each character in the specified WordArt in relation to the width of the character. Can be a value from zero through 5.
@property (copy) NSString *wordArtText;  // Returns or sets the text associated with this word art shape.

- (void) toggleVerticalText;  // Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.

@end

// Represents an word art object in the drawing layer.
@interface WordWordArt : WordShape

@property (readonly) BOOL bold;  // Returns if the text of the word art shape is formatted as bold.
@property (copy, readonly) NSString *fontName;  // Returns the font name of the font used by this word art shape.
@property (readonly) double fontSize;  // Returns the font size of the font used by this word art shape.
@property (readonly) BOOL italic;  // Returns if the text of the word art shape is formatted as italic.
@property (readonly) WordMsoPresetTextEffect presetWordArtEffect;  // Returns the style of the specified word art.
@property (copy, readonly) WordWordArtFormat *wordArtFormat;  // Returns the word art format object associated with the word art shape object.
@property (copy, readonly) NSString *wordArtText;  // The text associated with this word art shape.


@end

// Represents all the properties for wrapping text around a shape.
@interface WordWrapFormat : SBObject <WordGenericMethods>

@property BOOL allowOverlap;  // Returns or sets a value that specifies whether a given shape can overlap other shapes.
@property double distanceBottom;  // Returns or sets the distance in points between the document text and the bottom edge of the text-free area surrounding the specified shape.
@property double distanceLeft;  // Returns or sets the distance in points between the document text and the left edge of the text-free area surrounding the specified shape.
@property double distanceRight;  // Returns or sets the distance in points between the document text and the right edge of the text-free area surrounding the specified shape.
@property double distanceTop;  // Returns or sets the distance in points between the document text and the top edge of the text-free area surrounding the specified shape.
@property WordWdWrapSideType wrapSide;  // Returns or sets a value that indicates whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.
@property WordWdWrapType wrapType;  // Returns or sets the wrap type for the specified shape.


@end



/*
 * Text Suite
 */

// Represents a single built-in or user-defined style. The Word style object includes style attributes, font, font style, paragraph spacing, and so on, as properties of the Word style object.
@interface WordWordStyle : SBObject <WordGenericMethods>

@property BOOL automaticallyUpdate;  // True if the style is automatically redefined based on the selection. False if Word prompts for confirmation before redefining the style based on the selection.
@property WordWdBuiltinStyle baseStyle;  // Returns or sets an existing style on which you can base the formatting of another style.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property (readonly) BOOL builtIn;  // Returns true if the specified object is one of the built-in styles in Word.
@property (copy, readonly) NSString *objectDescription;  // Returns the description of the specified style. For example, a typical description for the Heading 2 style might be Normal, Font: Arial, 12 pt, Bold, Italic, Space Before 12 pt After 3 pt, KeepWithNext, Level 2.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this Word style object.
@property (copy, readonly) WordFrame *frame;  // Returns the frame object associated with this Word style object.
@property (readonly) BOOL inUse;  // Returns true if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.
@property WordWdLanguageID languageID;  // Returns or sets the language for the Word style object
@property WordWdLanguageID languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (readonly) NSInteger listLevelNumber;  // Returns the list level for the specified style.
@property (copy, readonly) WordListTemplate *listTemplate;  // Returns the list template object associated with this Word style object.
@property (copy) NSString *nameLocal;  // Returns or sets the name of a built-in style in the language of the user. Setting this property renames a user-defined style or adds an alias to a built-in style. 
@property WordWdBuiltinStyle nextParagraphStyle;  // Returns or sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the specified style.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores for this Word style
@property BOOL noSpaceBetweenSame;  // Returns or sets if Microsoft Word suppresses space between paragraphs of the same style
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this Word style object.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the Word style.
@property (readonly) WordWdStyleType styleType;  // Returns the style type.
@property (copy, readonly) WordTableStyle *tableStyle;  // Returns table style properties for this style.

- (void) linkToListTemplateListTemplate:(WordListTemplate *)listTemplate listLevelNumber:(NSInteger)listLevelNumber;  // Links the specified style to a list template so that the style's formatting can be applied to lists.

@end

// Represents all the formatting for a paragraph.
@interface WordParagraphFormat : SBObject <WordGenericMethods>

- (SBElementArray<WordTabStop *> *) tabStops;

@property BOOL addSpaceBetweenEastAsianAndAlpha;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs.
@property BOOL addSpaceBetweenEastAsianAndDigit;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs.
@property WordWdParagraphAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified paragraphs.
@property BOOL autoAdjustRightIndent;  // Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if youÔøïve specified a set number of characters per line.
@property WordWdBaselineAlignment baseLineAlignment;  // Returns or sets a constant that represents the vertical position of fonts on a line.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property double characterUnitFirstLineIndent;  // Returns or sets the value in characters for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property double characterUnitLeftIndent;  // Returns or sets the left indent value in characters for the specified paragraphs.
@property double characterUnitRightIndent;  // Returns or sets the right indent value in characters for the specified paragraphs.
@property BOOL disableLineHeightGrid;  // Returns or sets if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified.
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value in points for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property BOOL halfWidthPunctuationOnTopOfLine;  // Returns or sets if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs.
@property BOOL hangingPunctuation;  // Returns or sets if hanging punctuation is enabled for the specified paragraphs.
@property BOOL hyphenation;  // Returns or sets if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation.
@property BOOL keepTogether;  // Returns or sets if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
@property BOOL keepWithNext;  // Returns or sets if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
@property double lineSpacing;  // Returns or sets the line spacing in points for the specified paragraphs.
@property WordWdLineSpacing lineSpacingRule;  // Returns or sets the line spacing for the specified paragraphs.
@property double lineUnitAfter;  // Returns or sets the amount of spacing in gridlines after the specified paragraphs.
@property double lineUnitBefore;  // Returns or sets the amount of spacing in gridlines before the specified paragraphs.
@property BOOL noLineNumber;  // Returns or set if line numbers are repressed for the specified paragraphs.
@property WordWdOutlineLevel outlineLevel;  // Returns or sets the outline level for the specified paragraphs.
@property BOOL pageBreakBefore;  // Returns or sets if a page break is forced before the specified paragraphs.
@property double paragraphFormatLeftIndent;  // Returns or sets the left indent in points for the specified paragraphs.
@property double paragraphFormatRightIndent;  // Returns or sets the right indent in points for the specified paragraphs.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the paragraph format object.
@property double spaceAfter;  // Returns or sets the spacing in points after the specified paragraphs.
@property BOOL spaceAfterAuto;  // Returns or sets if Microsoft Word automatically sets the amount of spacing after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing in points before the specified paragraphs.
@property BOOL spaceBeforeAuto;  // Returns or sets if Microsoft Word automatically sets the amount of spacing before the specified paragraphs.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style associated with the replacement object.
@property BOOL widowControl;  // Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document.
@property BOOL wordWrap;  // Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames.


@end

// Represents a single paragraph in a selection, range, or document.
@interface WordParagraph : SBObject <WordGenericMethods>

- (SBElementArray<WordTabStop *> *) tabStops;

@property BOOL addSpaceBetweenEastAsianAndAlpha;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs.
@property BOOL addSpaceBetweenEastAsianAndDigit;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs.
@property WordWdParagraphAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified paragraphs.
@property BOOL autoAdjustRightIndent;  // Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if youÔøïve specified a set number of characters per line.
@property WordWdBaselineAlignment baseLineAlignment;  // Returns or sets a constant that represents the vertical position of fonts on a line.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property double characterUnitFirstLineIndent;  // Returns or sets the value in characters for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property double characterUnitLeftIndent;  // Returns or sets the left indent value in characters for the specified paragraphs.
@property double characterUnitRightIndent;  // Returns or sets the right indent value in characters for the specified paragraphs.
@property BOOL disableLineHeightGrid;  // Returns or sets if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified.
@property (copy, readonly) WordDropCap *dropCap;  // Returns the drop cap object associated with this paragraph object.
@property BOOL eastAsianLineBreakControl;
@property double firstLineIndent;  // Returns or sets the value in points for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
@property BOOL halfWidthPunctuationOnTopOfLine;  // Returns or sets if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs.
@property BOOL hangingPunctuation;  // Returns or sets if hanging punctuation is enabled for the specified paragraphs.
@property BOOL hyphenation;  // Returns or sets if the specified paragraphs are included in automatic hyphenation. False if the specified paragraphs are to be excluded from automatic hyphenation.
@property BOOL keepTogether;  // Returns or sets if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document.
@property BOOL keepWithNext;  // Returns or sets if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document.
@property double lineSpacing;  // Returns or sets the line spacing in points for the specified paragraphs.
@property WordWdLineSpacing lineSpacingRule;  // Returns or sets the line spacing for the specified paragraphs.
@property double lineUnitAfter;  // Returns or sets the amount of spacing in gridlines after the specified paragraphs.
@property double lineUnitBefore;  // Returns or sets the amount of spacing in gridlines before the specified paragraphs.
@property BOOL noLineNumber;  // Returns or set if line numbers are repressed for the specified paragraphs.
@property WordWdOutlineLevel outlineLevel;  // Returns or sets the outline level for the specified paragraphs.
@property BOOL pageBreakBefore;  // Returns or sets if a page break is forced before the specified paragraphs.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this paragraph object.
@property double paragraphLeftIndent;  // Returns or sets the left indent in points for the specified paragraphs.
@property NSInteger paragraph_id;  // Returns the paragraph id for the specified paragraphs.
@property double rightIndent;  // Returns or sets the right indent in points for the specified paragraphs.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the paragraph object.
@property double spaceAfter;  // Returns or sets the spacing in points after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing in points before the specified paragraphs.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style associated with the replacement object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the section object.
@property NSInteger text_id;  // Returns the text id for the specified paragraphs.
@property BOOL widowControl;  // Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document.
@property BOOL wordWrap;  // Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames.

- (void) indent;  // Indents one or more paragraphs by one level.
- (WordParagraph *) nextParagraph;  // Returns the next paragraph object.
- (void) outdent;  // Removes one level of indent for one or more paragraphs.
- (void) outlineDemote;  // Applies the next heading level style Heading 1 through Heading 8 to the specified paragraph or paragraphs. For example, if a paragraph is formatted with the Heading 2 style, this method demotes the paragraph by changing the style to Heading 3.
- (void) outlineDemoteToBody;  // Demotes the specified paragraph or paragraphs to body text by applying the Normal style.
- (void) outlinePromote;  // Applies the previous heading level style Heading 1 through Heading 8 to the specified paragraph or paragraphs. For example, if a paragraph is formatted with the Heading 2 style, this method promotes the paragraph by changing the style to Heading 1.
- (WordParagraph *) previousParagraph;  // Returns the previous paragraph object.

@end

// Represents a single section in a selection, range, or document.
@interface WordSection : SBObject <WordGenericMethods>

@property (copy, readonly) WordBorderOptions *borderOptions;
@property (copy) WordPageSetup *pageSetup;  // Returns or sets a page setup object associated with this section object
@property BOOL protectedForForms;  // Returns or sets if the specified section is protected for forms. When a section is protected for forms, you can select and modify text only in form fields.
@property (readonly) NSInteger sectionIndex;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the section object.

- (WordHeaderFooter *) getFooterIndex:(WordWdHeaderFooterIndex)index;  // Returns a specific footer object
- (WordHeaderFooter *) getHeaderIndex:(WordWdHeaderFooterIndex)index;  // Returns a specific header object

@end

// Contains shading attributes for an object.
@interface WordShading : SBObject <WordGenericMethods>

@property (copy) NSArray<NSNumber *> *backgroundPatternColor;  // Returns or sets the RGB color that's applied to the background of the shading object.
@property WordWdColorIndex backgroundPatternColorIndex;  // Returns or sets the color index that's applied to the background of the shading object.
@property WordMsoThemeColorIndex backgroundPatternColorThemeIndex;  // Returns or sets the color that's applied to the background of the shading object.
@property (copy) NSArray<NSNumber *> *foregroundPatternColor;  // Returns or sets the RGB color that's applied to the foreground of the shading object.
@property WordWdColorIndex foregroundPatternColorIndex;  // Returns or sets the color index that's applied to the foreground of the shading object.
@property WordMsoThemeColorIndex foregroundPatternColorThemeIndex;  // Returns or sets the color that's applied to the foreground of the shading object. This color is applied to the dots and lines in the shading pattern.
@property WordWdTextureIndex texture;  // Returns or sets the shading texture for the specified object.


@end

// Represents a contiguous area in a document. Each text range object is defined by a starting and ending character position.
@interface WordTextRange : SBObject <WordGenericMethods>

- (SBElementArray<WordCharacter *> *) characters;
- (SBElementArray<WordWord *> *) words;
- (SBElementArray<WordSentence *> *) sentences;
- (SBElementArray<WordTable *> *) tables;
- (SBElementArray<WordFootnote *> *) footnotes;
- (SBElementArray<WordEndnote *> *) endnotes;
- (SBElementArray<WordWordComment *> *) WordComments;
- (SBElementArray<WordCell *> *) cells;
- (SBElementArray<WordSection *> *) sections;
- (SBElementArray<WordParagraph *> *) paragraphs;
- (SBElementArray<WordField *> *) fields;
- (SBElementArray<WordFormField *> *) formFields;
- (SBElementArray<WordFrame *> *) frames;
- (SBElementArray<WordBookmark *> *) bookmarks;
- (SBElementArray<WordRevision *> *) revisions;
- (SBElementArray<WordHyperlinkObject *> *) hyperlinkObjects;
- (SBElementArray<WordSubdocument *> *) subdocuments;
- (SBElementArray<WordColumn *> *) columns;
- (SBElementArray<WordRow *> *) rows;
- (SBElementArray<WordShape *> *) shapes;
- (SBElementArray<WordReadabilityStatistic *> *) readabilityStatistics;
- (SBElementArray<WordGrammaticalError *> *) grammaticalErrors;
- (SBElementArray<WordSpellingError *> *) spellingErrors;
- (SBElementArray<WordInlineShape *> *) inlineShapes;
- (SBElementArray<WordMathObject *> *) mathObjects;
- (SBElementArray<WordCoauthLock *> *) coauthLocks;
- (SBElementArray<WordCoauthUpdate *> *) coauthUpdates;
- (SBElementArray<WordConflict *> *) conflicts;

@property BOOL bold;  // Returns or sets if the text associated with the text range is formatted as bold.
@property (readonly) NSInteger bookmarkId;  // Returns the number of the bookmark that encloses the beginning of the text range. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property WordWdCharacterCase character_case;  // Synonym for case
- (WordWdCharacterCase) character_case;  // Returns or sets a constant that represents the case of the text in the text range.
- (void) setCase: (WordWdCharacterCase) character_case;
@property (copy, readonly) WordColumnOptions *columnOptions;
@property BOOL combineCharacters;
@property (copy) NSString *content;  // Returns or sets the text in the text range.
@property BOOL disableCharacterSpaceGrid;  // Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding text range object.
@property WordWdEmphasisMark emphasisMark;  // Returns or sets the emphasis mark for a character or designated character string.
@property (readonly) NSInteger endOfContent;  // Returns the ending character position of the text range.
@property (copy, readonly) WordEndnoteOptions *endnoteOptions;  // Returns the endnote options object for this text range.
@property (copy, readonly) WordFieldOptions *fieldOptions;
@property (copy, readonly) WordFind *findObject;  // Returns the find object associated with this text range.
@property double fitTextWidth;  // Returns or sets the width in which Microsoft Word fits the text in the text range.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this text range.
@property (copy, readonly) WordFootnoteOptions *footnoteOptions;  // Returns the footnote options object for this text range.
@property (copy) WordTextRange *formattedText;  // Returns or sets a text range object that includes the formatted text in the text range.
@property BOOL grammarChecked;  // True if a grammar check has been run on the text range. False if some of the text range hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false.
@property WordWdColorIndex highlightColorIndex;  // Returns or sets the highlight color for the text range.
@property WordWdHorizontalInVerticalType horizontalInVertical;
@property (readonly) BOOL isEndOfRowMark;  // Returns true if the text range is collapsed and is located at the end-of-row mark in a table.
@property BOOL italic;  // Returns or sets if the text associated with the text range is formatted as italic.
@property WordWdLanguageID languageID;  // Returns or sets the language for the text range object
@property WordWdLanguageID languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (copy, readonly) WordListFormat *listFormat;  // Returns the list format object associated with this text range.
@property (copy, readonly) WordTextRange *nextStoryRange;  // Returns a text range object that refers to the next story
@property BOOL noProofing;  // Returns or sets if the spelling and grammar checker should ignore documents based on this text range.
@property WordWdTextOrientation orientation;  // Returns or sets the orientation of text in a text range when the text direction feature is enabled.
@property (copy) WordPageSetup *pageSetup;  // Returns or sets the page setup object associated with this text range.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this text range.
@property (readonly) NSInteger previousBookmarkId;  // Returns the number of the last bookmark that starts before or at the same place as the text range, It returns zero if there's no corresponding bookmark.
@property (copy, readonly) WordRangeEndnoteOptions *rangeEndnoteOptions;
@property (copy, readonly) WordRangeFootnoteOptions *rangeFootnoteOptions;
@property (copy, readonly) WordRowOptions *rowOptions;
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with this text range.
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are included in the elements of this text range.
@property BOOL spellingChecked;  // True if a spelling check has been run on the text range. False if some of the text range hasn't been checked for spelling. To see if the document contains spelling errors, get the count of spelling errors for the text range.
@property (readonly) NSInteger startOfContent;  // Returns the starting character position of the text range.
@property (readonly) NSInteger storyLength;  // Returns the number of characters in the story that contains the text range.
@property (readonly) WordWdStoryType storyType;  // Returns the story type for the text range.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style for this range.
@property BOOL subdocumentsExpanded;  // Returns or sets if the subdocuments in the specified text range are expanded.
@property WordWdLanguageID supplementalLanguageID;  // Returns or sets the language for the text range object
@property (copy) WordTextRetrievalMode *textRetrievalMode;  // Returns or sets the text retrieval object that controls how text is retrieved from this text range.
@property WordWdTwoLinesInOneType twoLinesInOne;  // Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.  Read/write
@property WordWdUnderline underline;  // Returns or sets the type of underline applied to the text range.

- (void) autoFormatTextRange;  // Automatically formats a text range.
- (double) calculateRange;  // Calculates a mathematical expression within a text range.
- (WordTextRange *) changeEndOfRangeBy:(WordWdUnits)by extendType:(WordWdMovementType)extendType;  // Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) changeStartOfRangeBy:(WordWdUnits)by extendType:(WordWdMovementType)extendType;  // Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.
- (BOOL) checkGrammar;  // Begins a grammar check for the text range.  Returns true if the text range had no errors
- (BOOL) checkSpellingCustomDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Begins a spelling check for the text range.  Returns true if the text range had no errors
- (void) checkSynonyms;  // Displays the thesaurus dialog box, which lists alternative word choices, or synonyms, for the text in the text range.
- (WordTextRange *) collapseRangeDirection:(WordWdCollapseDirection)direction;  // Collapses this text range to the starting or ending position and returns a new text range object. After a text range is collapsed, the starting and ending points are equal.
- (NSInteger) computeTextRangeStatisticsStatistic:(WordWdStatistic)statistic;  // Returns a statistic based on the contents of the specified text range.
- (WordTable *) convertToTableSeparator:(WordWdTableFieldSeparator)separator numberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns initialColumnWidth:(NSInteger)initialColumnWidth tableFormat:(WordWdTableFormat)tableFormat applyBorders:(BOOL)applyBorders applyShading:(BOOL)applyShading applyFont:(BOOL)applyFont applyColor:(BOOL)applyColor applyHeadingRows:(BOOL)applyHeadingRows applyLastRow:(BOOL)applyLastRow applyFirstColumn:(BOOL)applyFirstColumn applyLastColumn:(BOOL)applyLastColumn autoFit:(BOOL)autoFit autoFitBehavior:(WordWdAutoFitBehavior)autoFitBehavior defaultTableBehavior:(WordWdDefaultTableBehavior)defaultTableBehavior;  // Converts text within a text range to a table.
- (void) copyAsPicture NS_RETURNS_NOT_RETAINED;  // Copies the content of the text range as a picture.
- (void) copyObject NS_RETURNS_NOT_RETAINED;  // Copies the content of the text range to the clipboard.
- (void) cutObject;  // Removes the content of the text range from the document and places it on the clipboard.
- (WordTextRange *) expandBy:(WordWdUnits)by;  // Expands the specified range. This method returns the new range object or missing value if an error occurred.
- (NSString *) getRangeInformationInformationType:(WordWdInformation)informationType;  // Returns requested information about the text range. 
- (WordTextRange *) goToNextWhat:(WordWdGoToItem)what;  // Returns a text range object that refers to the start position of the next item or location specified by the what argument.
- (WordTextRange *) goToPreviousWhat:(WordWdGoToItem)what;  // Returns a text range object that refers to the start position of the previous item or location specified by the what argument.
- (BOOL) inRangeTextRange:(WordTextRange *)textRange;  // Returns true if the text range to which the method is applied is contained in the range specified by the text range argument.
- (BOOL) inStoryTextRange:(WordTextRange *)textRange;  // Returns true if the text range to which the method is applied is in the same story as the range specified by the text range argument.
- (BOOL) isEquivalentTextRange:(WordTextRange *)textRange;  // True if the selection or range to which this method is applied is equal to the range specified by the text range argument. This method compares the starting and ending character positions, as well as the story type.
- (void) modifyEnclosureEnclosureStyle:(WordWdEncloseStyle)enclosureStyle enclosureType:(WordWdEnclosureType)enclosureType enclosedText:(NSString *)enclosedText;  // Adds, modifies, or removes an enclosure around the specified character or characters.
- (WordTextRange *) moveEndOfRangeBy:(WordWdUnits)by count:(NSInteger)count;  // Moves the ending character position of the range.  This method returns the new text range object or missing value if an error occurred.
- (WordTextRange *) moveRangeBy:(WordWdUnits)by count:(NSInteger)count;  // Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeEndUntilCharacters:(NSString *)characters count:(WordWdRecoveryType)count;  // Moves the end position of the specified range until any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeEndWhileCharacters:(NSString *)characters count:(WordWdRecoveryType)count;  // Moves the ending character position of a the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeStartUntilCharacters:(NSString *)characters count:(WordWdRecoveryType)count;  // Moves the start position of the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeStartWhileCharacters:(NSString *)characters count:(WordWdRecoveryType)count;  // Moves the start position of the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeUntilCharacters:(NSString *)characters count:(WordWdRecoveryType)count;  // Moves the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeWhileCharacters:(NSString *)characters count:(WordWdRecoveryType)count;  // Moves the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveStartOfRangeBy:(WordWdUnits)by count:(NSInteger)count;  // Moves the starting character position of the range. This method returns the new text range object or missing value if an error occurred.
- (WordTextRange *) navigateTo:(WordWdGoToItem)to position:(WordWdGoToDirection)position count:(NSInteger)count name:(NSString *)name;  // Returns a text range object that represents the start position of the specified item, such as a page, bookmark, or field.
- (WordTextRange *) nextRangeBy:(WordWdUnits)by count:(NSInteger)count;  // Returns a text range object relative to the specified text range.
- (WordTextRange *) nextSubdocument;  // Returns a new text range object to the next subdocument. If there isn't another subdocument, an error occurs.
- (void) pasteAndFormatType:(WordWdRecoveryType)type;  // Pastes the selected table cells and formats them as specified.
- (void) pasteAppendTable;  // Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.
- (void) pasteAsNestedTable;  // Pastes a cell or group of cells as a nested table into the text range.
- (void) pasteExcelTableLinkedToExcel:(BOOL)linkedToExcel wordFormatting:(BOOL)wordFormatting RTF:(BOOL)RTF;  // Pastes and formats a Microsoft Excel table.
- (void) pasteObject;  // Inserts the contents of the clipboard at the specified text range.
- (void) pasteSpecialLink:(BOOL)link placement:(WordWdOLEPlacement)placement displayAsIcon:(BOOL)displayAsIcon dataType:(WordWdPasteDataType)dataType iconLabel:(NSString *)iconLabel;  // Inserts the contents of the clipboard. Unlike with the paste method, with paste special you can control the format of the pasted information and optionally establish a link to the source file - for example, a Microsoft Excel worksheet.
- (void) phoneticGuideText:(NSString *)text alignmentType:(WordWdPhoneticGuideAlignmentType)alignmentType raise:(NSInteger)raise fontSize:(NSInteger)fontSize fontName:(NSString *)fontName;
- (WordTextRange *) previousRangeBy:(WordWdUnits)by count:(NSInteger)count;  // Returns a text range object relative to the specified text range.
- (WordTextRange *) previousSubdocument;  // Returns a new text range object to the previous subdocument. If there isn't another subdocument, an error occurs.
- (void) relocateDirection:(WordWdRelocate)direction;  // In outline view, moves the paragraphs within the text range after the next visible paragraph or before the previous visible paragraph. Body text moves with a heading only if the body text is collapsed in outline view or if it's part of the text range.
- (WordTextRange *) setRangeStart:(NSInteger)start end:(NSInteger)end;  // Returns a text range object by using the specified starting and ending character positions.
- (void) sortExcludeHeader:(BOOL)excludeHeader fieldNumber:(NSInteger)fieldNumber sortFieldType:(WordWdSortFieldType)sortFieldType sortOrder:(WordWdSortOrder)sortOrder fieldNumberTwo:(NSInteger)fieldNumberTwo sortFieldTypeTwo:(WordWdSortFieldType)sortFieldTypeTwo sortOrderTwo:(WordWdSortOrder)sortOrderTwo fieldNumberThree:(NSInteger)fieldNumberThree sortFieldTypeThree:(WordWdSortFieldType)sortFieldTypeThree sortOrderThree:(WordWdSortOrder)sortOrderThree sortColumn:(BOOL)sortColumn separator:(WordWdSortSeparator)separator caseSensitive:(BOOL)caseSensitive languageID:(WordWdLanguageID)languageID;  // Sorts the paragraphs in the specified text range.
- (void) sortAscending;  // Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) sortDescending;  // Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (NSDictionary *) textRangeSpellingSuggestionsCustomDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary suggestionMode:(WordWdSpellingWordType)suggestionMode customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Returns a record that contains the spelling error type and the list of words suggested as replacements for the first word in the specified range

@end

@interface WordCharacter : WordTextRange


@end

@interface WordGrammaticalError : WordTextRange


@end

@interface WordSentence : WordTextRange


@end

@interface WordSpellingError : WordTextRange


@end

@interface WordStoryRange : WordTextRange


@end

@interface WordWord : WordTextRange


@end



/*
 * Proofing Suite
 */

// Represents a single autocorrect entry.
@interface WordAutocorrectEntry : SBObject <WordGenericMethods>

@property (copy) NSString *autocorrectValue;  // Returns or sets the value of the auto correct entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of the auto correct entry.
@property (readonly) BOOL richText;  // Returns if formatting is stored with the autocorrect entry replacement text.

- (void) applyCorrectionToRange:(WordTextRange *)toRange;  // Replaces a range with the value of the specified autocorrect entry.

@end

// Represents the autocorrect functionality in Word.
@interface WordAutocorrect : SBObject <WordGenericMethods>

- (SBElementArray<WordAutocorrectEntry *> *) autocorrectEntries;
- (SBElementArray<WordFirstLetterException *> *) firstLetterExceptions;
- (SBElementArray<WordTwoInitialCapsException *> *) twoInitialCapsExceptions;
- (SBElementArray<WordOtherCorrectionsException *> *) otherCorrectionsExceptions;

@property BOOL correctDays;  // Returns or sets if Word automatically capitalizes the first letter of days of the week.
@property BOOL correctInitialCaps;  // Returns or sets if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, WOrd is corrected to Word.
@property BOOL correctSentenceCaps;  // Returns or sets if Word automatically capitalizes the first letter in each sentence.
@property BOOL correctTableCaps;  // Returns or sets if Word automatically capitalizes the first letter in each table cell.
@property BOOL firstLetterAutoAdd;  // Returns or sets if Word automatically adds abbreviations to the list of autocorrect first letter exceptions.
@property BOOL otherCorrectionsAutoAdd;  // Returns or sets if Microsoft Word automatically adds words to the list of other autocorrect exceptions. Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct.
@property BOOL replaceText;  // Returns or sets if Microsoft Word automatically replaces specified text with entries from the autocorrect list.
@property BOOL replaceTextFromSpellingChecker;  // Returns or sets if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types.
@property BOOL showAutocorrectSmartButton;  // Returns or sets if Word shows the AutoCorrect smart button which allows you to review the correction when an automatic correction occurs.
@property BOOL twoInitialCapsAutoAdd;  // Returns or sets if Microsoft Word automatically adds words to the list of autocorrect initial caps exceptions.


@end

// Represents a dictionary.
@interface WordDictionary : SBObject <WordGenericMethods>

@property (readonly) WordWdDictionaryType dictionaryType;  // Returns the dictionary type.
@property WordWdLanguageID languageID;  // Returns or sets the language for the dictionary object
@property BOOL languageSpecific;  // Returns or sets if the custom dictionary is to be used only with text formatted for a specific language.
@property (copy, readonly) NSString *name;  // Returns or sets the name of this dictionary object.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified dictionary in HFS path style.
@property (readonly) BOOL readOnly;  // Returns true if the specified dictionary cannot be changed.


@end

// Represents an abbreviation excluded from automatic correction.
@interface WordFirstLetterException : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of this first letter exception object.


@end

// Represents a language used for proofing or formatting in Microsoft Word.
@interface WordLanguage : SBObject <WordGenericMethods>

@property (copy, readonly) WordDictionary *activeGrammarDictionary;  // Returns a dictionary object that represents the active grammar dictionary for the specified language
@property (copy, readonly) WordDictionary *activeHyphenationDictionary;  // Returns a dictionary object that represents the active hyphenation dictionary for the specified language.
@property (copy, readonly) WordDictionary *activeSpellingDictionary;  // Returns a dictionary object that represents the active spelling dictionary for the specified language.
@property (copy, readonly) WordDictionary *activeThesaurusDictionary;  // Returns a dictionary object that represents the active thesaurus dictionary for the specified language.
@property (copy) NSString *defaultWritingStyle;  // Returns or sets the default writing style used by the grammar checker for the specified language. The name of the writing style is the localized name for the specified language.
@property (readonly) WordWdLanguageID languageID;  // Returns an enumeration that identifies the specified language.
@property (copy, readonly) NSString *name;  // Returns the name of the language
@property (copy, readonly) NSString *nameLocal;  // Returns the name of a proofing tool language in the language of the user.
@property WordWdDictionaryType spellingDictionaryType;  // Returns or sets the proofing tool type
@property (copy, readonly) NSArray<NSString *> *writingStyleList;  // Returns a list of strings that contains the names of all writing styles available for the specified language.


@end

// Represents a single auto correct exception.
@interface WordOtherCorrectionsException : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of this other corrections exception object.


@end

// Represents one of the readability statistics for a document or range.
@interface WordReadabilityStatistic : SBObject <WordGenericMethods>

@property (copy, readonly) NSString *name;  // The name of this readability statistic object.
@property (readonly) double readabilityValue;  // The value of this readability statistic object.


@end

// Represents the information about synonyms, antonyms, related words, or related expressions for the specified range or a given string.
@interface WordSynonymInfo : SBObject <WordGenericMethods>

@property (copy, readonly) NSArray<NSString *> *antonyms;  // Returns a list of antonyms for the word or phrase.
@property (copy, readonly) NSString *context;  // Returns the word or phrase that was looked up by the thesaurus.
@property (readonly) BOOL found;  // Returns true if the thesaurus finds synonyms, antonyms, related words, or related expressions for the word or phrase.
@property (readonly) NSInteger meaningCount;  // Returns the number of entries in the list of meanings found in the thesaurus for the word or phrase. Returns zero if no meanings were found.
@property (copy, readonly) NSArray<NSString *> *meanings;  // Returns the list of meanings for the word or phrase.
@property (copy, readonly) NSArray<NSAppleEventDescriptor *> *partOfSpeech;  // Returns a list of the parts of speech corresponding to the meanings found for the word or phrase looked up in the thesaurus.
@property (copy, readonly) NSArray<NSString *> *relatedExpressions;  // Returns a list of expressions related to the specified word or phrase. 
@property (copy, readonly) NSArray<NSString *> *relatedWords;  // Returns a list of words related to the specified word or phrase.

- (NSArray<NSString *> *) getSynonymListForItemToCheck:(NSString *)itemToCheck;  // Get the list of synonyms for a particular word
- (NSArray<NSString *> *) getSynonymListFromMeaningIndex:(NSInteger)meaningIndex;  // Get the list of synonyms using an index into the list of meanings

@end

// Represents a single initial-capital autocorrect exception.
@interface WordTwoInitialCapsException : SBObject <WordGenericMethods>

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of this two initial caps exception object.


@end



/*
 * Table Suite
 */

// Represents a single table cell.
@interface WordCell : SBObject <WordGenericMethods>

- (SBElementArray<WordTable *> *) tables;

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordColumn *column;  // Returns the column object that contains this cell object.
@property (readonly) NSInteger columnIndex;  // Returns the number of the column that contains the specified cell.
@property BOOL fitText;  // Returns or sets if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width.
@property double height;  // Returns or sets the height of the object.
@property WordWdRowHeightRule heightRule;  // Returns or sets the rule for determining the height of the specified cells.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified cell.
@property (copy, readonly) WordCell *nextCell;  // Returns the next cell object
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified cell.
@property WordWdPreferredWidthType preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified column.
@property (copy, readonly) WordCell *previousCell;  // Returns the previous cell object
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordRow *row;  // Returns the row object that contains this cell object.
@property (readonly) NSInteger rowIndex;  // Returns the number of the row that contains the specified cell.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the cell object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the cell object.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property WordWdCellVerticalAlignment verticalAlignment;  // Returns or sets the vertical alignment of text in one or more cells of a table.
@property double width;  // Returns or sets the width of the object.
@property BOOL wordWrap;  // Returns or set  if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same.

- (void) autoSum;  // Inserts an = Formula field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.
- (void) formulaFormulaString:(NSString *)formulaString numberFormatString:(NSString *)numberFormatString;  // Inserts an = Formula field that contains the specified formula into a table cell.
- (void) mergeCellWith:(WordCell *)with;  // Merges the specified table cell with another cell. The result is a single table cell.
- (void) splitCellNumberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns;  // Splits a single table cell into multiple cells.

@end

// Represents options that can be set for columns.
@interface WordColumnOptions : SBObject <WordGenericMethods>

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double defaultWidth;  // Returns or sets the default width of columns.
@property (readonly) NSInteger nestingLevel;
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified columns. 
@property WordWdPreferredWidthType preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified columns. 
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with columns.

- (void) distributeWidth;  // Adjusts the width of the specified columns or cells so that they're equal.

@end

// Represents a single table column.
@interface WordColumn : SBObject <WordGenericMethods>

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this column object.
@property (readonly) NSInteger columnIndex;  // Returns the index for the position of the object in its container element list.
@property (readonly) BOOL isFirst;  // Returns if the specified column is the first one in the table.
@property (readonly) BOOL isLast;  // Returns if the specified column is the last one in the table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified column.
@property (copy, readonly) WordColumn *nextColumn;  // Returns the next column object
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified column.
@property WordWdPreferredWidthType preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified column.
@property (copy, readonly) WordColumn *previousColumn;  // Returns the previous column object
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the column object.
@property double width;  // Returns or sets the width of the object.


@end

@interface WordCondition : SBObject <WordGenericMethods>

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns the border options object associated with this condition.
@property double bottomPadding;
@property (copy, readonly) WordFont *fontObject;
@property double leftPadding;
@property (copy) WordParagraphFormat *paragraphFormat;
@property double rightPadding;
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with this condition.
@property double topPadding;


@end

// Represents options that can be set for rows.
@interface WordRowOptions : SBObject <WordGenericMethods>

@property WordWdRowAlignment alignment;  // Returns or sets a constant that represents the alignment for rows.
@property BOOL allowBreakAcrossPages;  // Returns or sets if the text in a table row or rows are allowed to split across a page break.
@property BOOL allowOverlap;  // Returns or sets a value that specifies whether the specified rows can overlap other rows.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double distanceBottom;  // Returns or sets the distance in points between the document text and the bottom edge of the specified table. This property doesn't have any effect if wrap around text is false. 
@property double distanceLeft;  // Returns or sets the distance in points between the document text and the left edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property double distanceRight;  // Returns or sets the distance in points between the document text and the right edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property double distanceTop;  // Returns or sets the distance in points between the document text and the top edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property BOOL headingFormat;  // Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page.
@property double height;  // Returns or sets the height of the object.
@property WordWdRowHeightRule heightRule;  // Returns or sets the rule for determining the height of the specified rows.
@property double horizontalPosition;  // Returns or sets the horizontal distance between the edge of the rows.
@property (readonly) NSInteger nestingLevel;
@property WordWdRelativeHorizontalPosition relativeHorizontalPosition;  // Specifies to what the horizontal position of a group of rows is relative.
@property WordWdRelativeVerticalPosition relativeVerticalPosition;  // Specifies to what the vertical position of a group of rows is relative.
@property double rowLeftIndent;  // Returns or sets the left indent in points for the specified rows.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with rows.
@property double spaceBetweenColumns;  // Returns or sets the distance in points between text in adjacent columns of the specified row or rows.
@property double verticalPosition;  // Returns or sets the vertical distance between the edge of the rows.
@property BOOL wrapAroundText;  // Returns or sets whether text should wrap around the specified rows. 

- (void) distributeRowHeight;  // Adjusts the height of the specified rows or cells so that they're equal.

@end

// Represents a row in a table.
@interface WordRow : SBObject <WordGenericMethods>

- (SBElementArray<WordCell *> *) cells;

@property WordWdRowAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified row.
@property BOOL allowBreakAcrossPages;  // Returns or sets if the text in a row or rows are allowed to split across a page break.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property BOOL headingFormat;  // Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page.
@property double height;  // Returns or sets the height of the object.
@property WordWdRowHeightRule heightRule;  // Returns or sets the rule for determining the height of the specified rows.
@property (readonly) BOOL isFirst;  // Returns if the specified row is the first one in the table.
@property (readonly) BOOL isLast;  // Returns if the specified row is the last one in the table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified row.
@property (copy, readonly) WordRow *nextRow;  // Returns the next row object
@property (copy, readonly) WordRow *previousRow;  // Returns the previous row object
@property (readonly) NSInteger rowIndex;  // Returns the index for the position of the object in its container element list.
@property double rowLeftIndent;  // Returns or sets the left indent in points for the specified rows.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the row object.
@property double spaceBetweenColumns;  // Returns or sets the distance in points between text in adjacent columns of the specified row or rows.  
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the row object.


@end

@interface WordTableStyle : SBObject <WordGenericMethods>

@property WordWdRowAlignment alignment;  // Returns or sets a constant that represents the alignment for the specified row.
@property BOOL allowBreakAccrossPage;
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property (copy, readonly) WordCondition *bottomLeftCellCondition;  // Returns the conditional style for the bottom left cell.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordCondition *bottomRightCellCondition;  // Returns the conditional style for the bottom right cell.
@property NSInteger columnStripe;
@property (copy, readonly) WordCondition *evenColumnCondition;  // Returns the conditional style for even column bands.
@property (copy, readonly) WordCondition *evenRowCondition;  // Returns the conditional style for even row bands.
@property (copy, readonly) WordCondition *firstColumnCondition;  // Returns the conditional style for the first column.
@property (copy, readonly) WordCondition *firstRowCondition;  // Returns the conditional style for the first row.
@property (copy, readonly) WordCondition *lastColumnCondition;  // Returns the conditional style for the last column.
@property (copy, readonly) WordCondition *lastRowCondition;  // Returns the conditional style for the last row.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordCondition *oddColumnCondition;  // Returns the conditional style for odd column bands.
@property (copy, readonly) WordCondition *oddRowCondition;  // Returns the conditional style for odd row bands.
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property double rowLeftIndent;  // Returns or sets the left indent in points for this table style.
@property NSInteger rowStripe;
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with this table style.
@property double spacing;  // Returns or sets the spacing between the cells in a table.
@property (copy, readonly) WordCondition *topLeftCellCondition;  // Returns the conditional style for the top left cell.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordCondition *topRightCellCondition;  // Returns the conditional style for the top right cell.


@end

// Represents a single table.
@interface WordTable : SBObject <WordGenericMethods>

- (SBElementArray<WordColumn *> *) columns;
- (SBElementArray<WordRow *> *) rows;
- (SBElementArray<WordTable *> *) tables;

@property BOOL allowAutoFit;  // Returns or sets if Microsoft Word will automatically resize cells in a table to fit their contents.
@property BOOL allowPageBreaks;  // Returns or sets if Microsoft Word allows to break the specified table across pages.
@property BOOL applyStyleFirstColumn;
@property BOOL applyStyleHeadingRows;
@property BOOL applyStyleLastColumn;
@property BOOL applyStyleLastRow;
@property (readonly) WordWdTableFormat autoFormatType;  // Returns the type of automatic formatting that's been applied to the specified table.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordColumnOptions *columnOptions;  // Returns the column options object associated with this table object.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified table.
@property (readonly) NSInteger numberOfColumns;  // Returns the number of columns in this table
@property (readonly) NSInteger numberOfRows;  // Returns the number of rows in this table
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified table.
@property WordWdPreferredWidthType preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified table. 
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordRowOptions *rowOptions;  // Returns the row options object associated with this table object.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the table object.
@property double spacing;  // Returns or sets the spacing between the cells in a table.
@property WordWdBuiltinStyle style;  // Returns or sets the Word style associated with the table object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the table object.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property (readonly) BOOL uniform;  // Returns if all the rows in a table have the same number of columns.

- (void) autoFitBehaviorBehavior:(WordWdAutoFitBehavior)behavior;  // Determines how Microsoft Word resizes a table when the autofit feature is used. Word can resize the table based on the content of the table cells or the width of the document window.
- (void) autoFormatTableTableFormat:(WordWdTableFormat)tableFormat applyBorders:(BOOL)applyBorders applyShading:(BOOL)applyShading applyFont:(BOOL)applyFont applyColor:(BOOL)applyColor applyHeadingRows:(BOOL)applyHeadingRows applyLastRow:(BOOL)applyLastRow applyFirstColumn:(BOOL)applyFirstColumn applyLastColumn:(BOOL)applyLastColumn autoFit:(BOOL)autoFit;  // Applies a predefined look to a table.
- (WordTextRange *) convertToTextSeparator:(WordWdTableFieldSeparator)separator nestedTables:(BOOL)nestedTables;  // Converts a table to text and returns a text range object that represents the delimited text.
- (WordCell *) getCellFromTableRow:(NSInteger)row column:(NSInteger)column;  // Returns a cell object that represents a cell in a table.
- (WordTable *) splitTableRow:(NSInteger)row;  // Inserts an empty paragraph immediately above the specified row in the table, and returns a table object that contains both the specified row and the rows that follow it.
- (void) updateAutoFormat;  // Updates the table with the characteristics of a predefined table format. For example, if you apply a table format with auto format and then insert rows and columns, the table may no longer match the predefined look.

@end

