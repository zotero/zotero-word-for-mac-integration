/*
 * Word.h
 */

#import <AppKit/AppKit.h>
#import <ScriptingBridge/ScriptingBridge.h>


@class WordBaseObject, WordBaseApplication, WordBaseDocument, WordBasicWindow, WordPrintSettings, WordCommandBarControl, WordCommandBarButton, WordCommandBarCombobox, WordCommandBarPopup, WordCommandBar, WordDocumentProperty, WordCustomDocumentProperty, WordWebPageFont, WordWordComment, WordWordList, WordWordOptions, WordAddIn, WordApplication, WordAutoTextEntry, WordBookmark, WordBorderOptions, WordBorder, WordBrowser, WordBuildingBlockCategory, WordBuildingBlockType, WordBuildingBlock, WordCaptionLabel, WordCheckBox, WordCoauthLock, WordCoauthUpdate, WordCoauthor, WordCoauthoring, WordConflict, WordCustomLabel, WordDataMergeDataField, WordDataMergeDataSource, WordDataMergeFieldName, WordDataMergeField, WordDataMerge, WordDefaultWebOptions, WordDialog, WordDocumentVersion, WordDocument, WordDropCap, WordDropDown, WordEndnoteOptions, WordEndnote, WordEnvelope, WordFieldOptions, WordField, WordFileConverter, WordFind, WordFont, WordFootnoteOptions, WordFootnote, WordFormField, WordFrame, WordHeaderFooter, WordHeadingStyle, WordHyperlinkObject, WordIndex, WordKeyBinding, WordLetterContent, WordLineNumbering, WordLinkFormat, WordListEntry, WordListFormat, WordListGallery, WordListLevel, WordListTemplate, WordMailingLabel, WordMathAccent, WordMathAutocorrectEntry, WordMathAutocorrect, WordMathBar, WordMathBorderBox, WordMathBox, WordMathBreak, WordMathDelimiter, WordMathEquationArray, WordMathFraction, WordMathFunc, WordMathFunction, WordMathGroupChar, WordMathLeftScripts, WordMathLowerLimit, WordMathMatrixColumn, WordMathMatrixRow, WordMathMatrix, WordMathNary, WordMathObject, WordMathPhantom, WordMathRadical, WordMathRecognizedFunction, WordMathSubAndSuperScript, WordMathSubscript, WordMathSuperscript, WordMathUpperLimit, WordPageNumberOptions, WordPageNumber, WordPageSetup, WordPane, WordRangeEndnoteOptions, WordRangeFootnoteOptions, WordRecentFile, WordReplacement, WordReviewer, WordRevision, WordSelectionObject, WordSubdocument, WordSystemObject, WordTabStop, WordTableOfAuthorities, WordTableOfContents, WordTableOfFigures, WordTemplate, WordTextColumn, WordTextInput, WordTextRetrievalMode, WordVariable, WordView, WordWebOptions, WordWindow, WordZoom, WordAdjustment, WordCalloutFormat, WordFillFormat, WordGlowFormat, WordHorizontalLineFormat, WordInlineShape, WordInlineHorizontalLine, WordInlinePictureBullet, WordInlinePicture, WordLineFormat, WordOfficeTheme, WordPictureFormat, WordReflectionFormat, WordShadowFormat, WordShape, WordCallout, WordLineShape, WordPicture, WordSoftEdgeFormat, WordStandardInlineHorizontalLine, WordTextBox, WordTextFrame, WordThemeColorScheme, WordThemeColor, WordThemeEffectScheme, WordThemeFontScheme, WordThemeFont, WordMajorThemeFont, WordMinorThemeFont, WordThemeFonts, WordThreeDFormat, WordWordArtFormat, WordWordArt, WordWrapFormat, WordWordStyle, WordParagraphFormat, WordParagraph, WordSection, WordShading, WordTextRange, WordCharacter, WordGrammaticalError, WordSentence, WordSpellingError, WordStoryRange, WordWord, WordAutocorrectEntry, WordAutocorrect, WordDictionary, WordFirstLetterException, WordLanguage, WordOtherCorrectionsException, WordReadabilityStatistic, WordSynonymInfo, WordTwoInitialCapsException, WordCell, WordColumnOptions, WordColumn, WordCondition, WordRowOptions, WordRow, WordTableStyle, WordTable;

enum WordSavo {
	WordSavoYes = 'yes ' /* Save objects now */,
	WordSavoNo = 'no  ' /* Do not save objects */,
	WordSavoAsk = 'ask ' /* Ask the user whether to save */
};
typedef enum WordSavo WordSavo;

enum WordKfrm {
	WordKfrmIndex = 'indx' /* keyform designating indexed access */,
	WordKfrmNamed = 'name' /* keyform designating named access */,
	WordKfrmId = 'ID  ' /* keyform designating access by unique identifier */
};
typedef enum WordKfrm WordKfrm;

enum WordEnum {
	WordEnumStandard = 'lwst' /* Standard PostScript error handling   */,
	WordEnumDetailed = 'lwdt' /* print a detailed report of Postscript errors */
};
typedef enum WordEnum WordEnum;

enum WordMlDs {
	WordMlDsLineDashStyleUnset = '\000\222\377\376',
	WordMlDsLineDashStyleSolid = '\000\223\000\001',
	WordMlDsLineDashStyleSquareDot = '\000\223\000\002',
	WordMlDsLineDashStyleRoundDot = '\000\223\000\003',
	WordMlDsLineDashStyleDash = '\000\223\000\004',
	WordMlDsLineDashStyleDashDot = '\000\223\000\005',
	WordMlDsLineDashStyleDashDotDot = '\000\223\000\006',
	WordMlDsLineDashStyleLongDash = '\000\223\000\007',
	WordMlDsLineDashStyleLongDashDot = '\000\223\000\010',
	WordMlDsLineDashStyleLongDashDotDot = '\000\223\000\011',
	WordMlDsLineDashStyleSystemDash = '\000\223\000\012',
	WordMlDsLineDashStyleSystemDot = '\000\223\000\013',
	WordMlDsLineDashStyleSystemDashDot = '\000\223\000\014'
};
typedef enum WordMlDs WordMlDs;

enum WordMLnS {
	WordMLnSLineStyleUnset = '\000\224\377\376',
	WordMLnSSingleLine = '\000\225\000\001',
	WordMLnSThinThinLine = '\000\225\000\002',
	WordMLnSThinThickLine = '\000\225\000\003',
	WordMLnSThickThinLine = '\000\225\000\004',
	WordMLnSThickBetweenThinLine = '\000\225\000\005'
};
typedef enum WordMLnS WordMLnS;

enum WordMAhS {
	WordMAhSArrowheadStyleUnset = '\000\221\377\376',
	WordMAhSNoArrowhead = '\000\222\000\001',
	WordMAhSTriangleArrowhead = '\000\222\000\002',
	WordMAhSOpen_arrowhead = '\000\222\000\003',
	WordMAhSStealthArrowhead = '\000\222\000\004',
	WordMAhSDiamondArrowhead = '\000\222\000\005',
	WordMAhSOvalArrowhead = '\000\222\000\006'
};
typedef enum WordMAhS WordMAhS;

enum WordMAhW {
	WordMAhWArrowheadWidthUnset = '\000\220\377\376',
	WordMAhWNarrowWidthArrowhead = '\000\221\000\001',
	WordMAhWMediumWidthArrowhead = '\000\221\000\002',
	WordMAhWWideArrowhead = '\000\221\000\003'
};
typedef enum WordMAhW WordMAhW;

enum WordMAhL {
	WordMAhLArrowheadLengthUnset = '\000\223\377\376',
	WordMAhLShortArrowhead = '\000\224\000\001',
	WordMAhLMediumArrowhead = '\000\224\000\002',
	WordMAhLLongArrowhead = '\000\224\000\003'
};
typedef enum WordMAhL WordMAhL;

enum WordMFdT {
	WordMFdTFillUnset = '\000c\377\376',
	WordMFdTFillSolid = '\000d\000\001',
	WordMFdTFillPatterned = '\000d\000\002',
	WordMFdTFillGradient = '\000d\000\003',
	WordMFdTFillTextured = '\000d\000\004',
	WordMFdTFillBackground = '\000d\000\005',
	WordMFdTFillPicture = '\000d\000\006'
};
typedef enum WordMFdT WordMFdT;

enum WordMGdS {
	WordMGdSGradientUnset = '\000d\377\376',
	WordMGdSHorizontalGradient = '\000e\000\001',
	WordMGdSVerticalGradient = '\000e\000\002',
	WordMGdSDiagonalUpGradient = '\000e\000\003',
	WordMGdSDiagonalDownGradient = '\000e\000\004',
	WordMGdSFromCornerGradient = '\000e\000\005',
	WordMGdSFromTitleGradient = '\000e\000\006',
	WordMGdSFromCenterGradient = '\000e\000\007'
};
typedef enum WordMGdS WordMGdS;

enum WordMGCt {
	WordMGCtGradientTypeUnset = '\003\357\377\376',
	WordMGCtSingleShadeGradientType = '\003\360\000\001',
	WordMGCtTwoColorsGradientType = '\003\360\000\002',
	WordMGCtPresetColorsGradientType = '\003\360\000\003',
	WordMGCtMultiColorsGradientType = '\003\360\000\004'
};
typedef enum WordMGCt WordMGCt;

enum WordMxtT {
	WordMxtTTextureTypeTextureTypeUnset = '\003\360\377\376',
	WordMxtTTextureTypePresetTexture = '\003\361\000\001',
	WordMxtTTextureTypeUserDefinedTexture = '\003\361\000\002'
};
typedef enum WordMxtT WordMxtT;

enum WordMPzT {
	WordMPzTPresetTextureUnset = '\000e\377\376',
	WordMPzTTexturePapyrus = '\000f\000\001',
	WordMPzTTextureCanvas = '\000f\000\002',
	WordMPzTTextureDenim = '\000f\000\003',
	WordMPzTTextureWovenMat = '\000f\000\004',
	WordMPzTTextureWaterDroplets = '\000f\000\005',
	WordMPzTTexturePaperBag = '\000f\000\006',
	WordMPzTTextureFishFossil = '\000f\000\007',
	WordMPzTTextureSand = '\000f\000\010',
	WordMPzTTextureGreenMarble = '\000f\000\011',
	WordMPzTTextureWhiteMarble = '\000f\000\012',
	WordMPzTTextureBrownMarble = '\000f\000\013',
	WordMPzTTextureGranite = '\000f\000\014',
	WordMPzTTextureNewsprint = '\000f\000\015',
	WordMPzTTextureRecycledPaper = '\000f\000\016',
	WordMPzTTextureParchment = '\000f\000\017',
	WordMPzTTextureStationery = '\000f\000\020',
	WordMPzTTextureBlueTissuePaper = '\000f\000\021',
	WordMPzTTexturePinkTissuePaper = '\000f\000\022',
	WordMPzTTexturePurpleMesh = '\000f\000\023',
	WordMPzTTextureBouquet = '\000f\000\024',
	WordMPzTTextureCork = '\000f\000\025',
	WordMPzTTextureWalnut = '\000f\000\026',
	WordMPzTTextureOak = '\000f\000\027',
	WordMPzTTextureMediumWood = '\000f\000\030'
};
typedef enum WordMPzT WordMPzT;

enum WordPpTy {
	WordPpTyPatternUnset = '\000f\377\376',
	WordPpTyFivePercentPattern = '\000g\000\001',
	WordPpTyTenPercentPattern = '\000g\000\002',
	WordPpTyTwentyPercentPattern = '\000g\000\003',
	WordPpTyTwentyFivePercentPattern = '\000g\000\004',
	WordPpTyThirtyPercentPattern = '\000g\000\005',
	WordPpTyFortyPercentPattern = '\000g\000\006',
	WordPpTyFiftyPercentPattern = '\000g\000\007',
	WordPpTySixtyPercentPattern = '\000g\000\010',
	WordPpTySeventyPercentPattern = '\000g\000\011',
	WordPpTySeventyFivePercentPattern = '\000g\000\012',
	WordPpTyEightyPercentPattern = '\000g\000\013',
	WordPpTyNinetyPercentPattern = '\000g\000\014',
	WordPpTyDarkHorizontalPattern = '\000g\000\015',
	WordPpTyDarkVerticalPattern = '\000g\000\016',
	WordPpTyDarkDownwardDiagonalPattern = '\000g\000\017',
	WordPpTyDarkUpwardDiagonalPattern = '\000g\000\020',
	WordPpTySmallCheckerBoardPattern = '\000g\000\021',
	WordPpTyTrellisPattern = '\000g\000\022',
	WordPpTyLightHorizontalPattern = '\000g\000\023',
	WordPpTyLightVerticalPattern = '\000g\000\024',
	WordPpTyLightDownwardDiagonalPattern = '\000g\000\025',
	WordPpTyLightUpwardDiagonalPattern = '\000g\000\026',
	WordPpTySmallGridPattern = '\000g\000\027',
	WordPpTyDottedDiamondPattern = '\000g\000\030',
	WordPpTyWideDownwardDiagonal = '\000g\000\031',
	WordPpTyWideUpwardDiagonalPattern = '\000g\000\032',
	WordPpTyDashedUpwardDiagonalPattern = '\000g\000\033',
	WordPpTyDashedDownwardDiagonalPattern = '\000g\000\034',
	WordPpTyNarrowVerticalPattern = '\000g\000\035',
	WordPpTyNarrowHorizontalPattern = '\000g\000\036',
	WordPpTyDashedVerticalPattern = '\000g\000\037',
	WordPpTyDashedHorizontalPattern = '\000g\000 ',
	WordPpTyLargeConfettiPattern = '\000g\000!',
	WordPpTyLargeGridPattern = '\000g\000\"',
	WordPpTyHorizontalBrickPattern = '\000g\000#',
	WordPpTyLargeCheckerBoardPattern = '\000g\000$',
	WordPpTySmallConfettiPattern = '\000g\000%',
	WordPpTyZigZagPattern = '\000g\000&',
	WordPpTySolidDiamondPattern = '\000g\000\'',
	WordPpTyDiagonalBrickPattern = '\000g\000(',
	WordPpTyOutlinedDiamondPattern = '\000g\000)',
	WordPpTyPlaidPattern = '\000g\000*',
	WordPpTySpherePattern = '\000g\000+',
	WordPpTyWeavePattern = '\000g\000,',
	WordPpTyDottedGridPattern = '\000g\000-',
	WordPpTyDivotPattern = '\000g\000.',
	WordPpTyShinglePattern = '\000g\000/',
	WordPpTyWavePattern = '\000g\0000',
	WordPpTyHorizontalPattern = '\000g\0001',
	WordPpTyVerticalPattern = '\000g\0002',
	WordPpTyCrossPattern = '\000g\0003',
	WordPpTyDownwardDiagonalPattern = '\000g\0004',
	WordPpTyUpwardDiagonalPattern = '\000g\0005',
	WordPpTyDiagonalCrossPattern = '\000g\0005'
};
typedef enum WordPpTy WordPpTy;

enum WordMPGb {
	WordMPGbPresetGradientUnset = '\000g\377\376',
	WordMPGbGradientEarlySunset = '\000h\000\001',
	WordMPGbGradientLateSunset = '\000h\000\002',
	WordMPGbGradientNightfall = '\000h\000\003',
	WordMPGbGradientDaybreak = '\000h\000\004',
	WordMPGbGradientHorizon = '\000h\000\005',
	WordMPGbGradientDesert = '\000h\000\006',
	WordMPGbGradientOcean = '\000h\000\007',
	WordMPGbGradientCalmWater = '\000h\000\010',
	WordMPGbGradientFire = '\000h\000\011',
	WordMPGbGradientFog = '\000h\000\012',
	WordMPGbGradientMoss = '\000h\000\013',
	WordMPGbGradientPeacock = '\000h\000\014',
	WordMPGbGradientWheat = '\000h\000\015',
	WordMPGbGradientParchment = '\000h\000\016',
	WordMPGbGradientMahogany = '\000h\000\017',
	WordMPGbGradientRainbow = '\000h\000\020',
	WordMPGbGradientRainbow2 = '\000h\000\021',
	WordMPGbGradientGold = '\000h\000\022',
	WordMPGbGradientGold2 = '\000h\000\023',
	WordMPGbGradientBrass = '\000h\000\024',
	WordMPGbGradientChrome = '\000h\000\025',
	WordMPGbGradientChrome2 = '\000h\000\026',
	WordMPGbGradientSilver = '\000h\000\027',
	WordMPGbGradientSapphire = '\000h\000\030'
};
typedef enum WordMPGb WordMPGb;

enum WordMSdT {
	WordMSdTShadowUnset = '\003_\377\376',
	WordMSdTShadow1 = '\003`\000\001',
	WordMSdTShadow2 = '\003`\000\002',
	WordMSdTShadow3 = '\003`\000\003',
	WordMSdTShadow4 = '\003`\000\004',
	WordMSdTShadow5 = '\003`\000\005',
	WordMSdTShadow6 = '\003`\000\006',
	WordMSdTShadow7 = '\003`\000\007',
	WordMSdTShadow8 = '\003`\000\010',
	WordMSdTShadow9 = '\003`\000\011',
	WordMSdTShadow10 = '\003`\000\012',
	WordMSdTShadow11 = '\003`\000\013',
	WordMSdTShadow12 = '\003`\000\014',
	WordMSdTShadow13 = '\003`\000\015',
	WordMSdTShadow14 = '\003`\000\016',
	WordMSdTShadow15 = '\003`\000\017',
	WordMSdTShadow16 = '\003`\000\020',
	WordMSdTShadow17 = '\003`\000\021',
	WordMSdTShadow18 = '\003`\000\022',
	WordMSdTShadow19 = '\003`\000\023',
	WordMSdTShadow20 = '\003`\000\024',
	WordMSdTShadow21 = '\003`\000\025',
	WordMSdTShadow22 = '\003`\000\026',
	WordMSdTShadow23 = '\003`\000\027',
	WordMSdTShadow24 = '\003`\000\030',
	WordMSdTShadow25 = '\003`\000\031',
	WordMSdTShadow26 = '\003`\000\032',
	WordMSdTShadow27 = '\003`\000\033',
	WordMSdTShadow28 = '\003`\000\034',
	WordMSdTShadow29 = '\003`\000\035',
	WordMSdTShadow30 = '\003`\000\036',
	WordMSdTShadow31 = '\003`\000\037',
	WordMSdTShadow32 = '\003`\000 ',
	WordMSdTShadow33 = '\003`\000!',
	WordMSdTShadow34 = '\003`\000\"',
	WordMSdTShadow35 = '\003`\000#',
	WordMSdTShadow36 = '\003`\000$',
	WordMSdTShadow37 = '\003`\000%',
	WordMSdTShadow38 = '\003`\000&',
	WordMSdTShadow39 = '\003`\000\'',
	WordMSdTShadow40 = '\003`\000(',
	WordMSdTShadow41 = '\003`\000)',
	WordMSdTShadow42 = '\003`\000*',
	WordMSdTShadow43 = '\003`\000+'
};
typedef enum WordMSdT WordMSdT;

enum WordMPXF {
	WordMPXFWordartFormatUnset = '\003\361\377\376',
	WordMPXFWordartFormat1 = '\003\362\000\000',
	WordMPXFWordartFormat2 = '\003\362\000\001',
	WordMPXFWordartFormat3 = '\003\362\000\002',
	WordMPXFWordartFormat4 = '\003\362\000\003',
	WordMPXFWordartFormat5 = '\003\362\000\004',
	WordMPXFWordartFormat6 = '\003\362\000\005',
	WordMPXFWordartFormat7 = '\003\362\000\006',
	WordMPXFWordartFormat8 = '\003\362\000\007',
	WordMPXFWordartFormat9 = '\003\362\000\010',
	WordMPXFWordartFormat10 = '\003\362\000\011',
	WordMPXFWordartFormat11 = '\003\362\000\012',
	WordMPXFWordartFormat12 = '\003\362\000\013',
	WordMPXFWordartFormat13 = '\003\362\000\014',
	WordMPXFWordartFormat14 = '\003\362\000\015',
	WordMPXFWordartFormat15 = '\003\362\000\016',
	WordMPXFWordartFormat16 = '\003\362\000\017',
	WordMPXFWordartFormat17 = '\003\362\000\020',
	WordMPXFWordartFormat18 = '\003\362\000\021',
	WordMPXFWordartFormat19 = '\003\362\000\022',
	WordMPXFWordartFormat20 = '\003\362\000\023',
	WordMPXFWordartFormat21 = '\003\362\000\024',
	WordMPXFWordartFormat22 = '\003\362\000\025',
	WordMPXFWordartFormat23 = '\003\362\000\026',
	WordMPXFWordartFormat24 = '\003\362\000\027',
	WordMPXFWordartFormat25 = '\003\362\000\030',
	WordMPXFWordartFormat26 = '\003\362\000\031',
	WordMPXFWordartFormat27 = '\003\362\000\032',
	WordMPXFWordartFormat28 = '\003\362\000\033',
	WordMPXFWordartFormat29 = '\003\362\000\034',
	WordMPXFWordartFormat30 = '\003\362\000\035'
};
typedef enum WordMPXF WordMPXF;

enum WordMPTs {
	WordMPTsTextEffectShapeUnset = '\000\227\377\376',
	WordMPTsPlainText = '\000\230\000\001',
	WordMPTsStop = '\000\230\000\002',
	WordMPTsTriangleUp = '\000\230\000\003',
	WordMPTsTriangleDown = '\000\230\000\004',
	WordMPTsChevronUp = '\000\230\000\005',
	WordMPTsChevronDown = '\000\230\000\006',
	WordMPTsRingInside = '\000\230\000\007',
	WordMPTsRingOutside = '\000\230\000\010',
	WordMPTsArchUpCurve = '\000\230\000\011',
	WordMPTsArchDownCurve = '\000\230\000\012',
	WordMPTsCircleCurve = '\000\230\000\013',
	WordMPTsButtonCurve = '\000\230\000\014',
	WordMPTsArchUpPour = '\000\230\000\015',
	WordMPTsArchDownPour = '\000\230\000\016',
	WordMPTsCirclePour = '\000\230\000\017',
	WordMPTsButtonPour = '\000\230\000\020',
	WordMPTsCurveUp = '\000\230\000\021',
	WordMPTsCurveDown = '\000\230\000\022',
	WordMPTsCanUp = '\000\230\000\023',
	WordMPTsCanDown = '\000\230\000\024',
	WordMPTsWave1 = '\000\230\000\025',
	WordMPTsWave2 = '\000\230\000\026',
	WordMPTsDoubleWave1 = '\000\230\000\027',
	WordMPTsDoubleWave2 = '\000\230\000\030',
	WordMPTsInflate = '\000\230\000\031',
	WordMPTsDeflate = '\000\230\000\032',
	WordMPTsInflateBottom = '\000\230\000\033',
	WordMPTsDeflateBottom = '\000\230\000\034',
	WordMPTsInflateTop = '\000\230\000\035',
	WordMPTsDeflateTop = '\000\230\000\036',
	WordMPTsDeflateInflate = '\000\230\000\037',
	WordMPTsDeflateInflateDeflate = '\000\230\000 ',
	WordMPTsFadeRight = '\000\230\000!',
	WordMPTsFadeLeft = '\000\230\000\"',
	WordMPTsFadeUp = '\000\230\000#',
	WordMPTsFadeDown = '\000\230\000$',
	WordMPTsSlantUp = '\000\230\000%',
	WordMPTsSlantDown = '\000\230\000&',
	WordMPTsCascadeUp = '\000\230\000\'',
	WordMPTsCascadeDown = '\000\230\000('
};
typedef enum WordMPTs WordMPTs;

enum WordMTxA {
	WordMTxATextEffectAlignmentUnset = '\000\226\377\376',
	WordMTxALeftTextEffectAlignment = '\000\227\000\001',
	WordMTxACenteredTextEffectAlignment = '\000\227\000\002',
	WordMTxARightTextEffectAlignment = '\000\227\000\003',
	WordMTxAJustifyTextEffectAlignment = '\000\227\000\004',
	WordMTxAWordJustifyTextEffectAlignment = '\000\227\000\005',
	WordMTxAStretchJustifyTextEffectAlignment = '\000\227\000\006'
};
typedef enum WordMTxA WordMTxA;

enum WordMPLd {
	WordMPLdPresetLightingDirectionUnset = '\000\233\377\376',
	WordMPLdLightFromTopLeft = '\000\234\000\001',
	WordMPLdLightFromTop = '\000\234\000\002',
	WordMPLdLightFromTopRight = '\000\234\000\003',
	WordMPLdLightFromLeft = '\000\234\000\004',
	WordMPLdLightFromNone = '\000\234\000\005',
	WordMPLdLightFromRight = '\000\234\000\006',
	WordMPLdLightFromBottomLeft = '\000\234\000\007',
	WordMPLdLightFromBottom = '\000\234\000\010',
	WordMPLdLightFromBottomRight = '\000\234\000\011'
};
typedef enum WordMPLd WordMPLd;

enum WordMlSf {
	WordMlSfLightingSoftnessUnset = '\000\234\377\376',
	WordMlSfLightingDim = '\000\235\000\001',
	WordMlSfLightingNormal = '\000\235\000\002',
	WordMlSfLightingBright = '\000\235\000\003'
};
typedef enum WordMlSf WordMlSf;

enum WordMPMt {
	WordMPMtPresetMaterialUnset = '\000\235\377\376',
	WordMPMtMatte = '\000\236\000\001',
	WordMPMtPlastic = '\000\236\000\002',
	WordMPMtMetal = '\000\236\000\003',
	WordMPMtWireframe = '\000\236\000\004',
	WordMPMtMatte2 = '\000\236\000\005',
	WordMPMtPlastic2 = '\000\236\000\006',
	WordMPMtMetal2 = '\000\236\000\007',
	WordMPMtWarmMatte = '\000\236\000\010',
	WordMPMtTranslucentPowder = '\000\236\000\011',
	WordMPMtPowder = '\000\236\000\012',
	WordMPMtDarkEdge = '\000\236\000\013',
	WordMPMtSoftEdge = '\000\236\000\014',
	WordMPMtMaterialClear = '\000\236\000\015',
	WordMPMtFlat = '\000\236\000\016',
	WordMPMtSoftMetal = '\000\236\000\017'
};
typedef enum WordMPMt WordMPMt;

enum WordMExD {
	WordMExDPresetExtrusionDirectionUnset = '\000\231\377\376',
	WordMExDExtrudeBottomRight = '\000\232\000\001',
	WordMExDExtrudeBottom = '\000\232\000\002',
	WordMExDExtrudeBottomLeft = '\000\232\000\003',
	WordMExDExtrudeRight = '\000\232\000\004',
	WordMExDExtrudeNone = '\000\232\000\005',
	WordMExDExtrudeLeft = '\000\232\000\006',
	WordMExDExtrudeTopRight = '\000\232\000\007',
	WordMExDExtrudeTop = '\000\232\000\010',
	WordMExDExtrudeTopLeft = '\000\232\000\011'
};
typedef enum WordMExD WordMExD;

enum WordM3DF {
	WordM3DFPresetThreeDFormatUnset = '\000\230\377\376',
	WordM3DFFormat1 = '\000\231\000\001',
	WordM3DFFormat2 = '\000\231\000\002',
	WordM3DFFormat3 = '\000\231\000\003',
	WordM3DFFormat4 = '\000\231\000\004',
	WordM3DFFormat5 = '\000\231\000\005',
	WordM3DFFormat6 = '\000\231\000\006',
	WordM3DFFormat7 = '\000\231\000\007',
	WordM3DFFormat8 = '\000\231\000\010',
	WordM3DFFormat9 = '\000\231\000\011',
	WordM3DFFormat10 = '\000\231\000\012',
	WordM3DFFormat11 = '\000\231\000\013',
	WordM3DFFormat12 = '\000\231\000\014',
	WordM3DFFormat13 = '\000\231\000\015',
	WordM3DFFormat14 = '\000\231\000\016',
	WordM3DFFormat15 = '\000\231\000\017',
	WordM3DFFormat16 = '\000\231\000\020',
	WordM3DFFormat17 = '\000\231\000\021',
	WordM3DFFormat18 = '\000\231\000\022',
	WordM3DFFormat19 = '\000\231\000\023',
	WordM3DFFormat20 = '\000\231\000\024'
};
typedef enum WordM3DF WordM3DF;

enum WordMExC {
	WordMExCExtrusionColorUnset = '\000\232\377\376',
	WordMExCExtrusionColorAutomatic = '\000\233\000\001',
	WordMExCExtrusionColorCustom = '\000\233\000\002'
};
typedef enum WordMExC WordMExC;

enum WordMCtT {
	WordMCtTConnectorTypeUnset = '\000h\377\376',
	WordMCtTStraight = '\000i\000\001',
	WordMCtTElbow = '\000i\000\002',
	WordMCtTCurve = '\000i\000\003'
};
typedef enum WordMCtT WordMCtT;

enum WordMHzA {
	WordMHzAHorizontalAnchorUnset = '\000\236\377\376',
	WordMHzAHorizontalAnchorNone = '\000\237\000\001',
	WordMHzAHorizontalAnchorCenter = '\000\237\000\002'
};
typedef enum WordMHzA WordMHzA;

enum WordMVtA {
	WordMVtAVerticalAnchorUnset = '\000\237\377\376',
	WordMVtAAnchorTop = '\000\240\000\001',
	WordMVtAAnchorTopBaseline = '\000\240\000\002',
	WordMVtAAnchorMiddle = '\000\240\000\003',
	WordMVtAAnchorBottom = '\000\240\000\004',
	WordMVtAAnchorBottomBaseline = '\000\240\000\005'
};
typedef enum WordMVtA WordMVtA;

enum WordMAsT {
	WordMAsTAutoshapeShapeTypeUnset = '\000i\377\376',
	WordMAsTAutoshapeRectangle = '\000j\000\001',
	WordMAsTAutoshapeParallelogram = '\000j\000\002',
	WordMAsTAutoshapeTrapezoid = '\000j\000\003',
	WordMAsTAutoshapeDiamond = '\000j\000\004',
	WordMAsTAutoshapeRoundedRectangle = '\000j\000\005',
	WordMAsTAutoshapeOctagon = '\000j\000\006',
	WordMAsTAutoshapeIsoscelesTriangle = '\000j\000\007',
	WordMAsTAutoshapeRightTriangle = '\000j\000\010',
	WordMAsTAutoshapeOval = '\000j\000\011',
	WordMAsTAutoshapeHexagon = '\000j\000\012',
	WordMAsTAutoshapeCross = '\000j\000\013',
	WordMAsTAutoshapeRegularPentagon = '\000j\000\014',
	WordMAsTAutoshapeCan = '\000j\000\015',
	WordMAsTAutoshapeCube = '\000j\000\016',
	WordMAsTAutoshapeBevel = '\000j\000\017',
	WordMAsTAutoshapeFoldedCorner = '\000j\000\020',
	WordMAsTAutoshapeSmileyFace = '\000j\000\021',
	WordMAsTAutoshapeDonut = '\000j\000\022',
	WordMAsTAutoshapeNoSymbol = '\000j\000\023',
	WordMAsTAutoshapeBlockArc = '\000j\000\024',
	WordMAsTAutoshapeHeart = '\000j\000\025',
	WordMAsTAutoshapeLightningBolt = '\000j\000\026',
	WordMAsTAutoshapeSun = '\000j\000\027',
	WordMAsTAutoshapeMoon = '\000j\000\030',
	WordMAsTAutoshapeArc = '\000j\000\031',
	WordMAsTAutoshapeDoubleBracket = '\000j\000\032',
	WordMAsTAutoshapeDoubleBrace = '\000j\000\033',
	WordMAsTAutoshapePlaque = '\000j\000\034',
	WordMAsTAutoshapeLeftBracket = '\000j\000\035',
	WordMAsTAutoshapeRightBracket = '\000j\000\036',
	WordMAsTAutoshapeLeftBrace = '\000j\000\037',
	WordMAsTAutoshapeRightBrace = '\000j\000 ',
	WordMAsTAutoshapeRightArrow = '\000j\000!',
	WordMAsTAutoshapeLeftArrow = '\000j\000\"',
	WordMAsTAutoshapeUpArrow = '\000j\000#',
	WordMAsTAutoshapeDownArrow = '\000j\000$',
	WordMAsTAutoshapeLeftRightArrow = '\000j\000%',
	WordMAsTAutoshapeUpDownArrow = '\000j\000&',
	WordMAsTAutoshapeQuadArrow = '\000j\000\'',
	WordMAsTAutoshapeLeftRightUpArrow = '\000j\000(',
	WordMAsTAutoshapeBentArrow = '\000j\000)',
	WordMAsTAutoshapeUTurnArrow = '\000j\000*',
	WordMAsTAutoshapeLeftUpArrow = '\000j\000+',
	WordMAsTAutoshapeBentUpArrow = '\000j\000,',
	WordMAsTAutoshapeCurvedRightArrow = '\000j\000-',
	WordMAsTAutoshapeCurvedLeftArrow = '\000j\000.',
	WordMAsTAutoshapeCurvedUpArrow = '\000j\000/',
	WordMAsTAutoshapeCurvedDownArrow = '\000j\0000',
	WordMAsTAutoshapeStripedRightArrow = '\000j\0001',
	WordMAsTAutoshapeNotchedRightArrow = '\000j\0002',
	WordMAsTAutoshapePentagon = '\000j\0003',
	WordMAsTAutoshapeChevron = '\000j\0004',
	WordMAsTAutoshapeRightArrowCallout = '\000j\0005',
	WordMAsTAutoshapeLeftArrowCallout = '\000j\0006',
	WordMAsTAutoshapeUpArrowCallout = '\000j\0007',
	WordMAsTAutoshapeDownArrowCallout = '\000j\0008',
	WordMAsTAutoshapeLeftRightArrowCallout = '\000j\0009',
	WordMAsTAutoshapeUpDownArrowCallout = '\000j\000:',
	WordMAsTAutoshapeQuadArrowCallout = '\000j\000;',
	WordMAsTAutoshapeCircularArrow = '\000j\000<',
	WordMAsTAutoshapeFlowchartProcess = '\000j\000=',
	WordMAsTAutoshapeFlowchartAlternateProcess = '\000j\000>',
	WordMAsTAutoshapeFlowchartDecision = '\000j\000\?',
	WordMAsTAutoshapeFlowchartData = '\000j\000@',
	WordMAsTAutoshapeFlowchartPredefinedProcess = '\000j\000A',
	WordMAsTAutoshapeFlowchartInternalStorage = '\000j\000B',
	WordMAsTAutoshapeFlowchartDocument = '\000j\000C',
	WordMAsTAutoshapeFlowchartMultiDocument = '\000j\000D',
	WordMAsTAutoshapeFlowchartTerminator = '\000j\000E',
	WordMAsTAutoshapeFlowchartPreparation = '\000j\000F',
	WordMAsTAutoshapeFlowchartManualInput = '\000j\000G',
	WordMAsTAutoshapeFlowchartManualOperation = '\000j\000H',
	WordMAsTAutoshapeFlowchartConnector = '\000j\000I',
	WordMAsTAutoshapeFlowchartOffpageConnector = '\000j\000J',
	WordMAsTAutoshapeFlowchartCard = '\000j\000K',
	WordMAsTAutoshapeFlowchartPunchedTape = '\000j\000L',
	WordMAsTAutoshapeFlowchartSummingJunction = '\000j\000M',
	WordMAsTAutoshapeFlowchartOr = '\000j\000N',
	WordMAsTAutoshapeFlowchartCollate = '\000j\000O',
	WordMAsTAutoshapeFlowchartSort = '\000j\000P',
	WordMAsTAutoshapeFlowchartExtract = '\000j\000Q',
	WordMAsTAutoshapeFlowchartMerge = '\000j\000R',
	WordMAsTAutoshapeFlowchartStoredData = '\000j\000S',
	WordMAsTAutoshapeFlowchartDelay = '\000j\000T',
	WordMAsTAutoshapeFlowchartSequentialAccessStorage = '\000j\000U',
	WordMAsTAutoshapeFlowchartMagneticDisk = '\000j\000V',
	WordMAsTAutoshapeFlowchartDirectAccessStorage = '\000j\000W',
	WordMAsTAutoshapeFlowchartDisplay = '\000j\000X',
	WordMAsTAutoshapeExplosionOne = '\000j\000Y',
	WordMAsTAutoshapeExplosionTwo = '\000j\000Z',
	WordMAsTAutoshapeFourPointStar = '\000j\000[',
	WordMAsTAutoshapeFivePointStar = '\000j\000\\',
	WordMAsTAutoshapeEightPointStar = '\000j\000]',
	WordMAsTAutoshapeSixteenPointStar = '\000j\000^',
	WordMAsTAutoshapeTwentyFourPointStar = '\000j\000_',
	WordMAsTAutoshapeThirtyTwoPointStar = '\000j\000`',
	WordMAsTAutoshapeUpRibbon = '\000j\000a',
	WordMAsTAutoshapeDownRibbon = '\000j\000b',
	WordMAsTAutoshapeCurvedUpRibbon = '\000j\000c',
	WordMAsTAutoshapeCurvedDownRibbon = '\000j\000d',
	WordMAsTAutoshapeVerticalScroll = '\000j\000e',
	WordMAsTAutoshapeHorizontalScroll = '\000j\000f',
	WordMAsTAutoshapeWave = '\000j\000g',
	WordMAsTAutoshapeDoubleWave = '\000j\000h',
	WordMAsTAutoshapeRectangularCallout = '\000j\000i',
	WordMAsTAutoshapeRoundedRectangularCallout = '\000j\000j',
	WordMAsTAutoshapeOvalCallout = '\000j\000k',
	WordMAsTAutoshapeCloudCallout = '\000j\000l',
	WordMAsTAutoshapeLineCalloutOne = '\000j\000m',
	WordMAsTAutoshapeLineCalloutTwo = '\000j\000n',
	WordMAsTAutoshapeLineCalloutThree = '\000j\000o',
	WordMAsTAutoshapeLineCalloutFour = '\000j\000p',
	WordMAsTAutoshapeLineCalloutOneAccentBar = '\000j\000q',
	WordMAsTAutoshapeLineCalloutTwoAccentBar = '\000j\000r',
	WordMAsTAutoshapeLineCalloutThreeAccentBar = '\000j\000s',
	WordMAsTAutoshapeLineCalloutFourAccentBar = '\000j\000t',
	WordMAsTAutoshapeLineCalloutOneNoBorder = '\000j\000u',
	WordMAsTAutoshapeLineCalloutTwoNoBorder = '\000j\000v',
	WordMAsTAutoshapeLineCalloutThreeNoBorder = '\000j\000w',
	WordMAsTAutoshapeLineCalloutFourNoBorder = '\000j\000x',
	WordMAsTAutoshapeCalloutOneBorderAndAccentBar = '\000j\000y',
	WordMAsTAutoshapeCalloutTwoBorderAndAccentBar = '\000j\000z',
	WordMAsTAutoshapeCalloutThreeBorderAndAccentBar = '\000j\000{',
	WordMAsTAutoshapeCalloutFourBorderAndAccentBar = '\000j\000|',
	WordMAsTAutoshapeActionButtonCustom = '\000j\000}',
	WordMAsTAutoshapeActionButtonHome = '\000j\000~',
	WordMAsTAutoshapeActionButtonHelp = '\000j\000\177',
	WordMAsTAutoshapeActionButtonInformation = '\000j\000\200',
	WordMAsTAutoshapeActionButtonBackOrPrevious = '\000j\000\201',
	WordMAsTAutoshapeActionButtonForwardOrNext = '\000j\000\202',
	WordMAsTAutoshapeActionButtonBeginning = '\000j\000\203',
	WordMAsTAutoshapeActionButtonEnd = '\000j\000\204',
	WordMAsTAutoshapeActionButtonReturn = '\000j\000\205',
	WordMAsTAutoshapeActionButtonDocument = '\000j\000\206',
	WordMAsTAutoshapeActionButtonSound = '\000j\000\207',
	WordMAsTAutoshapeActionButtonMovie = '\000j\000\210',
	WordMAsTAutoshapeBalloon = '\000j\000\211',
	WordMAsTAutoshapeNotPrimitive = '\000j\000\212',
	WordMAsTAutoshapeFlowchartOfflineStorage = '\000j\000\213',
	WordMAsTAutoshapeLeftRightRibbon = '\000j\000\214',
	WordMAsTAutoshapeDiagonalStripe = '\000j\000\215',
	WordMAsTAutoshapePie = '\000j\000\216',
	WordMAsTAutoshapeNonIsoscelesTrapezoid = '\000j\000\217',
	WordMAsTAutoshapeDecagon = '\000j\000\220',
	WordMAsTAutoshapeHeptagon = '\000j\000\221',
	WordMAsTAutoshapeDodecagon = '\000j\000\222',
	WordMAsTAutoshapeSixPointsStar = '\000j\000\223',
	WordMAsTAutoshapeSevenPointsStar = '\000j\000\224',
	WordMAsTAutoshapeTenPointsStar = '\000j\000\225',
	WordMAsTAutoshapeTwelvePointsStar = '\000j\000\226',
	WordMAsTAutoshapeRoundOneRectangle = '\000j\000\227',
	WordMAsTAutoshapeRoundTwoSameRectangle = '\000j\000\230',
	WordMAsTAutoshapeRoundTwoDiagonalRectangle = '\000j\000\231',
	WordMAsTAutoshapeSnipRoundRectangle = '\000j\000\232',
	WordMAsTAutoshapeSnipOneRectangle = '\000j\000\233',
	WordMAsTAutoshapeSnipTwoSameRectangle = '\000j\000\234',
	WordMAsTAutoshapeSnipTwoDiagonalRectangle = '\000j\000\235',
	WordMAsTAutoshapeFrame = '\000j\000\236',
	WordMAsTAutoshapeHalfFrame = '\000j\000\237',
	WordMAsTAutoshapeTear = '\000j\000\240',
	WordMAsTAutoshapeChord = '\000j\000\241',
	WordMAsTAutoshapeCorner = '\000j\000\242',
	WordMAsTAutoshapeMathPlus = '\000j\000\243',
	WordMAsTAutoshapeMathMinus = '\000j\000\244',
	WordMAsTAutoshapeMathMultiply = '\000j\000\245',
	WordMAsTAutoshapeMathDivide = '\000j\000\246',
	WordMAsTAutoshapeMathEqual = '\000j\000\247',
	WordMAsTAutoshapeMathNotEqual = '\000j\000\250',
	WordMAsTAutoshapeCornerTabs = '\000j\000\251',
	WordMAsTAutoshapeSquareTabs = '\000j\000\252',
	WordMAsTAutoshapePlaqueTabs = '\000j\000\253',
	WordMAsTAutoshapeGearSix = '\000j\000\254',
	WordMAsTAutoshapeGearNine = '\000j\000\255',
	WordMAsTAutoshapeFunnel = '\000j\000\256',
	WordMAsTAutoshapePieWedge = '\000j\000\257',
	WordMAsTAutoshapeLeftCircularArrow = '\000j\000\260',
	WordMAsTAutoshapeLeftRightCircularArrow = '\000j\000\261',
	WordMAsTAutoshapeSwooshArrow = '\000j\000\262',
	WordMAsTAutoshapeCloud = '\000j\000\263',
	WordMAsTAutoshapeChartX = '\000j\000\264',
	WordMAsTAutoshapeChartStar = '\000j\000\265',
	WordMAsTAutoshapeChartPlus = '\000j\000\266',
	WordMAsTAutoshapeLineInverse = '\000j\000\267'
};
typedef enum WordMAsT WordMAsT;

enum WordMShp {
	WordMShpShapeTypeUnset = '\000\213\377\376',
	WordMShpShapeTypeAuto = '\000\214\000\001',
	WordMShpShapeTypeCallout = '\000\214\000\002',
	WordMShpShapeTypeChart = '\000\214\000\003',
	WordMShpShapeTypeComment = '\000\214\000\004',
	WordMShpShapeTypeFreeForm = '\000\214\000\005',
	WordMShpShapeTypeGroup = '\000\214\000\006',
	WordMShpShapeTypeEmbeddedOLEControl = '\000\214\000\007',
	WordMShpShapeTypeFormControl = '\000\214\000\010',
	WordMShpShapeTypeLine = '\000\214\000\011',
	WordMShpShapeTypeLinkedOLEObject = '\000\214\000\012',
	WordMShpShapeTypeLinkedPicture = '\000\214\000\013',
	WordMShpShapeTypeOLEControl = '\000\214\000\014',
	WordMShpShapeTypePicture = '\000\214\000\015',
	WordMShpShapeTypePlaceHolder = '\000\214\000\016',
	WordMShpShapeTypeWordArt = '\000\214\000\017',
	WordMShpShapeTypeMedia = '\000\214\000\020',
	WordMShpShapeTypeTextBox = '\000\214\000\021',
	WordMShpShapeTypeTable = '\000\214\000\022',
	WordMShpShapeTypeCanvas = '\000\214\000\023',
	WordMShpShapeTypeDiagram = '\000\214\000\024',
	WordMShpShapeTypeInk = '\000\214\000\025',
	WordMShpShapeTypeInkComment = '\000\214\000\026',
	WordMShpShapeTypeSmartartGraphic = '\000\214\000\027',
	WordMShpShapeTypeSlicer = '\000\214\000\030'
};
typedef enum WordMShp WordMShp;

enum WordMCrT {
	WordMCrTColorTypeUnset = '\000j\377\376',
	WordMCrTRGB = '\000k\000\001',
	WordMCrTScheme = '\000k\000\002'
};
typedef enum WordMCrT WordMCrT;

enum WordMPc {
	WordMPcPictureColorTypeUnset = '\000\265\377\376',
	WordMPcPictureColorAutomatic = '\000\266\000\001',
	WordMPcPictureColorGrayScale = '\000\266\000\002',
	WordMPcPictureColorBlackAndWhite = '\000\266\000\003',
	WordMPcPictureColorWatermark = '\000\266\000\004'
};
typedef enum WordMPc WordMPc;

enum WordMCAt {
	WordMCAtAngleUnset = '\000k\377\376',
	WordMCAtAngleAutomatic = '\000l\000\001',
	WordMCAtAngle30 = '\000l\000\002',
	WordMCAtAngle45 = '\000l\000\003',
	WordMCAtAngle60 = '\000l\000\004',
	WordMCAtAngle90 = '\000l\000\005'
};
typedef enum WordMCAt WordMCAt;

enum WordMCDt {
	WordMCDtDropUnset = '\000l\377\376',
	WordMCDtDropCustom = '\000m\000\001',
	WordMCDtDropTop = '\000m\000\002',
	WordMCDtDropCenter = '\000m\000\003',
	WordMCDtDropBottom = '\000m\000\004'
};
typedef enum WordMCDt WordMCDt;

enum WordMCot {
	WordMCotCalloutUnset = '\000m\377\376',
	WordMCotCalloutOne = '\000n\000\001',
	WordMCotCalloutTwo = '\000n\000\002',
	WordMCotCalloutThree = '\000n\000\003',
	WordMCotCalloutFour = '\000n\000\004'
};
typedef enum WordMCot WordMCot;

enum WordTxOr {
	WordTxOrTextOrientationUnset = '\000\215\377\376',
	WordTxOrHorizontal = '\000\216\000\001',
	WordTxOrUpward = '\000\216\000\002',
	WordTxOrDownward = '\000\216\000\003',
	WordTxOrVerticalEastAsian = '\000\216\000\004',
	WordTxOrVertical = '\000\216\000\005',
	WordTxOrHorizontalRotatedEastAsian = '\000\216\000\006'
};
typedef enum WordTxOr WordTxOr;

enum WordMSFr {
	WordMSFrScaleFromTopLeft = '\000o\000\000',
	WordMSFrScaleFromMiddle = '\000o\000\001',
	WordMSFrScaleFromBottomRight = '\000o\000\002'
};
typedef enum WordMSFr WordMSFr;

enum WordMPzC {
	WordMPzCPresetCameraUnset = '\000\256\377\376',
	WordMPzCCameraLegacyObliqueFromTopLeft = '\000\257\000\001',
	WordMPzCCameraLegacyObliqueFromTop = '\000\257\000\002',
	WordMPzCCameraLegacyObliqueFromTopright = '\000\257\000\003',
	WordMPzCCameraLegacyObliqueFromLeft = '\000\257\000\004',
	WordMPzCCameraLegacyObliqueFromFront = '\000\257\000\005',
	WordMPzCCameraLegacyObliqueFromRight = '\000\257\000\006',
	WordMPzCCameraLegacyObliqueFromBottomLeft = '\000\257\000\007',
	WordMPzCCameraLegacyObliqueFromBottom = '\000\257\000\010',
	WordMPzCCameraLegacyObliqueFromBottomRight = '\000\257\000\011',
	WordMPzCCameraLegacyPerspectiveFromTopLeft = '\000\257\000\012',
	WordMPzCCameraLegacyPerspectiveFromTop = '\000\257\000\013',
	WordMPzCCameraLegacyPerspectiveFromTopRight = '\000\257\000\014',
	WordMPzCCameraLegacyPerspectiveFromLeft = '\000\257\000\015',
	WordMPzCCameraLegacyPerspectiveFromFront = '\000\257\000\016',
	WordMPzCCameraLegacyPerspectiveFromRight = '\000\257\000\017',
	WordMPzCCameraLegacyPerspectiveFromBottomLeft = '\000\257\000\020',
	WordMPzCCameraLegacyPerspectiveFromBottom = '\000\257\000\021',
	WordMPzCCameraLegacyPerspectiveFromBottomRight = '\000\257\000\022',
	WordMPzCCameraOrthographic = '\000\257\000\023',
	WordMPzCCameraIsometricFromTopUp = '\000\257\000\024',
	WordMPzCCameraIsometricFromTopDown = '\000\257\000\025',
	WordMPzCCameraIsometricFromBottomUp = '\000\257\000\026',
	WordMPzCCameraIsometricFromBottomDown = '\000\257\000\027',
	WordMPzCCameraIsometricFromLeftUp = '\000\257\000\030',
	WordMPzCCameraIsometricFromLeftDown = '\000\257\000\031',
	WordMPzCCameraIsometricFromRightUp = '\000\257\000\032',
	WordMPzCCameraIsometricFromRightDown = '\000\257\000\033',
	WordMPzCCameraIsometricOffAxis1FromLeft = '\000\257\000\034',
	WordMPzCCameraIsometricOffAxis1FromRight = '\000\257\000\035',
	WordMPzCCameraIsometricOffAxis1FromTop = '\000\257\000\036',
	WordMPzCCameraIsometricOffAxis2FromLeft = '\000\257\000\037',
	WordMPzCCameraIsometricOffAxis2FromRight = '\000\257\000 ',
	WordMPzCCameraIsometricOffAxis2FromTop = '\000\257\000!',
	WordMPzCCameraIsometricOffAxis3FromLeft = '\000\257\000\"',
	WordMPzCCameraIsometricOffAxis3FromRight = '\000\257\000#',
	WordMPzCCameraIsometricOffAxis3FromBottom = '\000\257\000$',
	WordMPzCCameraIsometricOffAxis4FromLeft = '\000\257\000%',
	WordMPzCCameraIsometricOffAxis4FromRight = '\000\257\000&',
	WordMPzCCameraIsometricOffAxis4FromBottom = '\000\257\000\'',
	WordMPzCCameraObliqueFromTopLeft = '\000\257\000(',
	WordMPzCCameraObliqueFromTop = '\000\257\000)',
	WordMPzCCameraObliqueFromTopRight = '\000\257\000*',
	WordMPzCCameraObliqueFromLeft = '\000\257\000+',
	WordMPzCCameraObliqueFromRight = '\000\257\000,',
	WordMPzCCameraObliqueFromBottomLeft = '\000\257\000-',
	WordMPzCCameraObliqueFromBottom = '\000\257\000.',
	WordMPzCCameraObliqueFromBottomRight = '\000\257\000/',
	WordMPzCCameraPerspectiveFromFront = '\000\257\0000',
	WordMPzCCameraPerspectiveFromLeft = '\000\257\0001',
	WordMPzCCameraPerspectiveFromRight = '\000\257\0002',
	WordMPzCCameraPerspectiveFromAbove = '\000\257\0003',
	WordMPzCCameraPerspectiveFromBelow = '\000\257\0004',
	WordMPzCCameraPerspectiveFromAboveFacingLeft = '\000\257\0005',
	WordMPzCCameraPerspectiveFromAboveFacingRight = '\000\257\0006',
	WordMPzCCameraPerspectiveContrastingFacingLeft = '\000\257\0007',
	WordMPzCCameraPerspectiveContrastingFacingRight = '\000\257\0008',
	WordMPzCCameraPerspectiveHeroicFacingLeft = '\000\257\0009',
	WordMPzCCameraPerspectiveHeroicFacingRight = '\000\257\000:',
	WordMPzCCameraPerspectiveHeroicExtremeFacingLeft = '\000\257\000;',
	WordMPzCCameraPerspectiveHeroicExtremeFacingRight = '\000\257\000<',
	WordMPzCCameraPerspectiveRelaxed = '\000\257\000=',
	WordMPzCCameraPerspectiveRelaxedModerately = '\000\257\000>'
};
typedef enum WordMPzC WordMPzC;

enum WordMLtT {
	WordMLtTLightRigUnset = '\000\257\377\376',
	WordMLtTLightRigFlat1 = '\000\260\000\001',
	WordMLtTLightRigFlat2 = '\000\260\000\002',
	WordMLtTLightRigFlat3 = '\000\260\000\003',
	WordMLtTLightRigFlat4 = '\000\260\000\004',
	WordMLtTLightRigNormal1 = '\000\260\000\005',
	WordMLtTLightRigNormal2 = '\000\260\000\006',
	WordMLtTLightRigNormal3 = '\000\260\000\007',
	WordMLtTLightRigNormal4 = '\000\260\000\010',
	WordMLtTLightRigHarsh1 = '\000\260\000\011',
	WordMLtTLightRigHarsh2 = '\000\260\000\012',
	WordMLtTLightRigHarsh3 = '\000\260\000\013',
	WordMLtTLightRigHarsh4 = '\000\260\000\014',
	WordMLtTLightRigThreePoint = '\000\260\000\015',
	WordMLtTLightRigBalanced = '\000\260\000\016',
	WordMLtTLightRigSoft = '\000\260\000\017',
	WordMLtTLightRigHarsh = '\000\260\000\020',
	WordMLtTLightRigFlood = '\000\260\000\021',
	WordMLtTLightRigContrasting = '\000\260\000\022',
	WordMLtTLightRigMorning = '\000\260\000\023',
	WordMLtTLightRigSunrise = '\000\260\000\024',
	WordMLtTLightRigSunset = '\000\260\000\025',
	WordMLtTLightRigChilly = '\000\260\000\026',
	WordMLtTLightRigFreezing = '\000\260\000\027',
	WordMLtTLightRigFlat = '\000\260\000\030',
	WordMLtTLightRigTwoPoint = '\000\260\000\031',
	WordMLtTLightRigGlow = '\000\260\000\032',
	WordMLtTLightRigBrightRoom = '\000\260\000\033'
};
typedef enum WordMLtT WordMLtT;

enum WordMBlT {
	WordMBlTBevelTypeUnset = '\000\260\377\376',
	WordMBlTBevelNone = '\000\261\000\001',
	WordMBlTBevelRelaxedInset = '\000\261\000\002',
	WordMBlTBevelCircle = '\000\261\000\003',
	WordMBlTBevelSlope = '\000\261\000\004',
	WordMBlTBevelCross = '\000\261\000\005',
	WordMBlTBevelAngle = '\000\261\000\006',
	WordMBlTBevelSoftRound = '\000\261\000\007',
	WordMBlTBevelConvex = '\000\261\000\010',
	WordMBlTBevelCoolSlant = '\000\261\000\011',
	WordMBlTBevelDivot = '\000\261\000\012',
	WordMBlTBevelRiblet = '\000\261\000\013',
	WordMBlTBevelHardEdge = '\000\261\000\014',
	WordMBlTBevelArtDeco = '\000\261\000\015'
};
typedef enum WordMBlT WordMBlT;

enum WordMSSt {
	WordMSStShadowStyleUnset = '\000\261\377\376',
	WordMSStShadowStyleInner = '\000\262\000\001',
	WordMSStShadowStyleOuter = '\000\262\000\002'
};
typedef enum WordMSSt WordMSSt;

enum WordPpgA {
	WordPpgAParagraphAlignmentUnset = '\000\346\377\376',
	WordPpgAParagraphAlignLeft = '\000\347\000\000',
	WordPpgAParagraphAlignCenter = '\000\347\000\001',
	WordPpgAParagraphAlignRight = '\000\347\000\002',
	WordPpgAParagraphAlignJustify = '\000\347\000\003',
	WordPpgAParagraphAlignDistribute = '\000\347\000\004',
	WordPpgAParagraphAlignThai = '\000\347\000\005',
	WordPpgAParagraphAlignJustifyLow = '\000\347\000\006'
};
typedef enum WordPpgA WordPpgA;

enum WordMTSt {
	WordMTStStrikeUnset = '\000\263\377\376',
	WordMTStNoStrike = '\000\264\000\000',
	WordMTStSingleStrike = '\000\264\000\001',
	WordMTStDoubleStrike = '\000\264\000\002'
};
typedef enum WordMTSt WordMTSt;

enum WordMTCa {
	WordMTCaCapsUnset = '\000\264\377\376',
	WordMTCaNoCaps = '\000\265\000\000',
	WordMTCaSmallCaps = '\000\265\000\001',
	WordMTCaAllCaps = '\000\265\000\002'
};
typedef enum WordMTCa WordMTCa;

enum WordMTUl {
	WordMTUlUnderlineUnset = '\003\356\377\376',
	WordMTUlNoUnderline = '\003\357\000\000',
	WordMTUlUnderlineWordsOnly = '\003\357\000\001',
	WordMTUlUnderlineSingleLine = '\003\357\000\002',
	WordMTUlUnderlineDoubleLine = '\003\357\000\003',
	WordMTUlUnderlineHeavyLine = '\003\357\000\004',
	WordMTUlUnderlineDottedLine = '\003\357\000\005',
	WordMTUlUnderlineHeavyDottedLine = '\003\357\000\006',
	WordMTUlUnderlineDashLine = '\003\357\000\007',
	WordMTUlUnderlineHeavyDashLine = '\003\357\000\010',
	WordMTUlUnderlineLongDashLine = '\003\357\000\011',
	WordMTUlUnderlineHeavyLongDashLine = '\003\357\000\012',
	WordMTUlUnderlineDotDashLine = '\003\357\000\013',
	WordMTUlUnderlineHeavyDotDashLine = '\003\357\000\014',
	WordMTUlUnderlineDotDotDashLine = '\003\357\000\015',
	WordMTUlUnderlineHeavyDotDotDashLine = '\003\357\000\016',
	WordMTUlUnderlineWavyLine = '\003\357\000\017',
	WordMTUlUnderlineHeavyWavyLine = '\003\357\000\020',
	WordMTUlUnderlineWavyDoubleLine = '\003\357\000\021'
};
typedef enum WordMTUl WordMTUl;

enum WordMTTA {
	WordMTTATabUnset = '\000\266\377\376',
	WordMTTALeftTab = '\000\267\000\000',
	WordMTTACenterTab = '\000\267\000\001',
	WordMTTARightTab = '\000\267\000\002',
	WordMTTADecimalTab = '\000\267\000\003'
};
typedef enum WordMTTA WordMTTA;

enum WordMTCW {
	WordMTCWCharacterWrapUnset = '\000\267\377\376',
	WordMTCWNoCharacterWrap = '\000\270\000\000',
	WordMTCWStandardCharacterWrap = '\000\270\000\001',
	WordMTCWStrictCharacterWrap = '\000\270\000\002',
	WordMTCWCustomCharacterWrap = '\000\270\000\003'
};
typedef enum WordMTCW WordMTCW;

enum WordMTFA {
	WordMTFAFontAlignmentUnset = '\000\270\377\376',
	WordMTFAAutomaticAlignment = '\000\271\000\000',
	WordMTFATopAlignment = '\000\271\000\001',
	WordMTFACenterAlignment = '\000\271\000\002',
	WordMTFABaselineAlignment = '\000\271\000\003',
	WordMTFABottomAlignment = '\000\271\000\004'
};
typedef enum WordMTFA WordMTFA;

enum WordPAtS {
	WordPAtSAutoSizeUnset = '\000\344\377\376',
	WordPAtSAutoSizeNone = '\000\345\000\000',
	WordPAtSShapeToFitText = '\000\345\000\001',
	WordPAtSTextToFitShape = '\000\345\000\002'
};
typedef enum WordPAtS WordPAtS;

enum WordMPFo {
	WordMPFoPathTypeUnset = '\000\272\377\376',
	WordMPFoNoPathType = '\000\273\000\000',
	WordMPFoPathType1 = '\000\273\000\001',
	WordMPFoPathType2 = '\000\273\000\002',
	WordMPFoPathType3 = '\000\273\000\003',
	WordMPFoPathType4 = '\000\273\000\004'
};
typedef enum WordMPFo WordMPFo;

enum WordMWFo {
	WordMWFoWarpFormatUnset = '\000\273\377\376',
	WordMWFoWarpFormat1 = '\000\274\000\000',
	WordMWFoWarpFormat2 = '\000\274\000\001',
	WordMWFoWarpFormat3 = '\000\274\000\002',
	WordMWFoWarpFormat4 = '\000\274\000\003',
	WordMWFoWarpFormat5 = '\000\274\000\004',
	WordMWFoWarpFormat6 = '\000\274\000\005',
	WordMWFoWarpFormat7 = '\000\274\000\006',
	WordMWFoWarpFormat8 = '\000\274\000\007',
	WordMWFoWarpFormat9 = '\000\274\000\010',
	WordMWFoWarpFormat10 = '\000\274\000\011',
	WordMWFoWarpFormat11 = '\000\274\000\012',
	WordMWFoWarpFormat12 = '\000\274\000\013',
	WordMWFoWarpFormat13 = '\000\274\000\014',
	WordMWFoWarpFormat14 = '\000\274\000\015',
	WordMWFoWarpFormat15 = '\000\274\000\016',
	WordMWFoWarpFormat16 = '\000\274\000\017',
	WordMWFoWarpFormat17 = '\000\274\000\020',
	WordMWFoWarpFormat18 = '\000\274\000\021',
	WordMWFoWarpFormat19 = '\000\274\000\022',
	WordMWFoWarpFormat20 = '\000\274\000\023',
	WordMWFoWarpFormat21 = '\000\274\000\024',
	WordMWFoWarpFormat22 = '\000\274\000\025',
	WordMWFoWarpFormat23 = '\000\274\000\026',
	WordMWFoWarpFormat24 = '\000\274\000\027',
	WordMWFoWarpFormat25 = '\000\274\000\030',
	WordMWFoWarpFormat26 = '\000\274\000\031',
	WordMWFoWarpFormat27 = '\000\274\000\032',
	WordMWFoWarpFormat28 = '\000\274\000\033',
	WordMWFoWarpFormat29 = '\000\274\000\034',
	WordMWFoWarpFormat30 = '\000\274\000\035',
	WordMWFoWarpFormat31 = '\000\274\000\036',
	WordMWFoWarpFormat32 = '\000\274\000\037',
	WordMWFoWarpFormat33 = '\000\274\000 ',
	WordMWFoWarpFormat34 = '\000\274\000!',
	WordMWFoWarpFormat35 = '\000\274\000\"',
	WordMWFoWarpFormat36 = '\000\274\000#'
};
typedef enum WordMWFo WordMWFo;

enum WordPCgC {
	WordPCgCCaseSentence = '\000\344\000\001',
	WordPCgCCaseLower = '\000\344\000\002',
	WordPCgCCaseUpper = '\000\344\000\003',
	WordPCgCCaseTitle = '\000\344\000\004',
	WordPCgCCaseToggle = '\000\344\000\005'
};
typedef enum WordPCgC WordPCgC;

enum WordMDTF {
	WordMDTFDateTimeFormatUnset = '\000\275\377\376',
	WordMDTFDateTimeFormatMdyy = '\000\276\000\001',
	WordMDTFDateTimeFormatDdddMMMMddyyyy = '\000\276\000\002',
	WordMDTFDateTimeFormatDMMMMyyyy = '\000\276\000\003',
	WordMDTFDateTimeFormatMMMMdyyyy = '\000\276\000\004',
	WordMDTFDateTimeFormatDMMMyy = '\000\276\000\005',
	WordMDTFDateTimeFormatMMMMyy = '\000\276\000\006',
	WordMDTFDateTimeFormatMMyy = '\000\276\000\007',
	WordMDTFDateTimeFormatMMddyyHmm = '\000\276\000\010',
	WordMDTFDateTimeFormatMMddyyhmmAMPM = '\000\276\000\011',
	WordMDTFDateTimeFormatHmm = '\000\276\000\012',
	WordMDTFDateTimeFormatHmmss = '\000\276\000\013',
	WordMDTFDateTimeFormatHmmAMPM = '\000\276\000\014',
	WordMDTFDateTimeFormatHmmssAMPM = '\000\276\000\015',
	WordMDTFDateTimeFormatFigureOut = '\000\276\000\016'
};
typedef enum WordMDTF WordMDTF;

enum WordMSET {
	WordMSETSoftEdgeUnset = '\000\277\377\376',
	WordMSETNoSoftEdge = '\000\300\000\000',
	WordMSETSoftEdgeType1 = '\000\300\000\001',
	WordMSETSoftEdgeType2 = '\000\300\000\002',
	WordMSETSoftEdgeType3 = '\000\300\000\003',
	WordMSETSoftEdgeType4 = '\000\300\000\004',
	WordMSETSoftEdgeType5 = '\000\300\000\005',
	WordMSETSoftEdgeType6 = '\000\300\000\006'
};
typedef enum WordMSET WordMSET;

enum WordMCSI {
	WordMCSIFirstDarkSchemeColor = '\000\301\000\001',
	WordMCSIFirstLightSchemeColor = '\000\301\000\002',
	WordMCSISecondDarkSchemeColor = '\000\301\000\003',
	WordMCSISecondLightSchemeColor = '\000\301\000\004',
	WordMCSIFirstAccentSchemeColor = '\000\301\000\005',
	WordMCSISecondAccentSchemeColor = '\000\301\000\006',
	WordMCSIThirdAccentSchemeColor = '\000\301\000\007',
	WordMCSIFourthAccentSchemeColor = '\000\301\000\010',
	WordMCSIFifthAccentSchemeColor = '\000\301\000\011',
	WordMCSISixthAccentSchemeColor = '\000\301\000\012',
	WordMCSIHyperlinkSchemeColor = '\000\301\000\013',
	WordMCSIFollowedHyperlinkSchemeColor = '\000\301\000\014'
};
typedef enum WordMCSI WordMCSI;

enum WordMCoI {
	WordMCoIThemeColorUnset = '\000\301\377\376',
	WordMCoINoThemeColor = '\000\302\000\000',
	WordMCoIFirstDarkThemeColor = '\000\302\000\001',
	WordMCoIFirstLightThemeColor = '\000\302\000\002',
	WordMCoISecondDarkThemeColor = '\000\302\000\003',
	WordMCoISecondLightThemeColor = '\000\302\000\004',
	WordMCoIFirstAccentThemeColor = '\000\302\000\005',
	WordMCoISecondAccentThemeColor = '\000\302\000\006',
	WordMCoIThirdAccentThemeColor = '\000\302\000\007',
	WordMCoIFourthAccentThemeColor = '\000\302\000\010',
	WordMCoIFifthAccentThemeColor = '\000\302\000\011',
	WordMCoISixthAccentThemeColor = '\000\302\000\012',
	WordMCoIHyperlinkThemeColor = '\000\302\000\013',
	WordMCoIFollowedHyperlinkThemeColor = '\000\302\000\014',
	WordMCoIFirstTextThemeColor = '\000\302\000\015',
	WordMCoIFirstBackgroundThemeColor = '\000\302\000\016',
	WordMCoISecondTextThemeColor = '\000\302\000\017',
	WordMCoISecondBackgroundThemeColor = '\000\302\000\020'
};
typedef enum WordMCoI WordMCoI;

enum WordMFLI {
	WordMFLIThemeFontLatin = '\000\303\000\001',
	WordMFLIThemeFontComplexScript = '\000\303\000\002',
	WordMFLIThemeFontHighAnsi = '\000\303\000\003',
	WordMFLIThemeFontEastAsian = '\000\303\000\004'
};
typedef enum WordMFLI WordMFLI;

enum WordMSSI {
	WordMSSIShapeStyleUnset = '\000\303\377\376',
	WordMSSIShapeNotAPreset = '\000\304\000\000',
	WordMSSIShapePreset1 = '\000\304\000\001',
	WordMSSIShapePreset2 = '\000\304\000\002',
	WordMSSIShapePreset3 = '\000\304\000\003',
	WordMSSIShapePreset4 = '\000\304\000\004',
	WordMSSIShapePreset5 = '\000\304\000\005',
	WordMSSIShapePreset6 = '\000\304\000\006',
	WordMSSIShapePreset7 = '\000\304\000\007',
	WordMSSIShapePreset8 = '\000\304\000\010',
	WordMSSIShapePreset9 = '\000\304\000\011',
	WordMSSIShapePreset10 = '\000\304\000\012',
	WordMSSIShapePreset11 = '\000\304\000\013',
	WordMSSIShapePreset12 = '\000\304\000\014',
	WordMSSIShapePreset13 = '\000\304\000\015',
	WordMSSIShapePreset14 = '\000\304\000\016',
	WordMSSIShapePreset15 = '\000\304\000\017',
	WordMSSIShapePreset16 = '\000\304\000\020',
	WordMSSIShapePreset17 = '\000\304\000\021',
	WordMSSIShapePreset18 = '\000\304\000\022',
	WordMSSIShapePreset19 = '\000\304\000\023',
	WordMSSIShapePreset20 = '\000\304\000\024',
	WordMSSIShapePreset21 = '\000\304\000\025',
	WordMSSIShapePreset22 = '\000\304\000\026',
	WordMSSIShapePreset23 = '\000\304\000\027',
	WordMSSIShapePreset24 = '\000\304\000\030',
	WordMSSIShapePreset25 = '\000\304\000\031',
	WordMSSIShapePreset26 = '\000\304\000\032',
	WordMSSIShapePreset27 = '\000\304\000\033',
	WordMSSIShapePreset28 = '\000\304\000\034',
	WordMSSIShapePreset29 = '\000\304\000\035',
	WordMSSIShapePreset30 = '\000\304\000\036',
	WordMSSIShapePreset31 = '\000\304\000\037',
	WordMSSIShapePreset32 = '\000\304\000 ',
	WordMSSIShapePreset33 = '\000\304\000!',
	WordMSSIShapePreset34 = '\000\304\000\"',
	WordMSSIShapePreset35 = '\000\304\000#',
	WordMSSIShapePreset36 = '\000\304\000$',
	WordMSSIShapePreset37 = '\000\304\000%',
	WordMSSIShapePreset38 = '\000\304\000&',
	WordMSSIShapePreset39 = '\000\304\000\'',
	WordMSSIShapePreset40 = '\000\304\000(',
	WordMSSIShapePreset41 = '\000\304\000)',
	WordMSSIShapePreset42 = '\000\304\000*',
	WordMSSILinePreset1 = '\000\304\'\021',
	WordMSSILinePreset2 = '\000\304\'\022',
	WordMSSILinePreset3 = '\000\304\'\023',
	WordMSSILinePreset4 = '\000\304\'\024',
	WordMSSILinePreset5 = '\000\304\'\025',
	WordMSSILinePreset6 = '\000\304\'\026',
	WordMSSILinePreset7 = '\000\304\'\027',
	WordMSSILinePreset8 = '\000\304\'\030',
	WordMSSILinePreset9 = '\000\304\'\031',
	WordMSSILinePreset10 = '\000\304\'\032',
	WordMSSILinePreset11 = '\000\304\'\033',
	WordMSSILinePreset12 = '\000\304\'\034',
	WordMSSILinePreset13 = '\000\304\'\035',
	WordMSSILinePreset14 = '\000\304\'\036',
	WordMSSILinePreset15 = '\000\304\'\037',
	WordMSSILinePreset16 = '\000\304\' ',
	WordMSSILinePreset17 = '\000\304\'!',
	WordMSSILinePreset18 = '\000\304\'\"',
	WordMSSILinePreset19 = '\000\304\'#',
	WordMSSILinePreset20 = '\000\304\'$',
	WordMSSILinePreset21 = '\000\304\'%'
};
typedef enum WordMSSI WordMSSI;

enum WordMBSI {
	WordMBSIBackgroundUnset = '\000\304\377\376',
	WordMBSIBackgroundNotAPreset = '\000\305\000\000',
	WordMBSIBackgroundPreset1 = '\000\305\000\001',
	WordMBSIBackgroundPreset2 = '\000\305\000\002',
	WordMBSIBackgroundPreset3 = '\000\305\000\003',
	WordMBSIBackgroundPreset4 = '\000\305\000\004',
	WordMBSIBackgroundPreset5 = '\000\305\000\005',
	WordMBSIBackgroundPreset6 = '\000\305\000\006',
	WordMBSIBackgroundPreset7 = '\000\305\000\007',
	WordMBSIBackgroundPreset8 = '\000\305\000\010',
	WordMBSIBackgroundPreset9 = '\000\305\000\011',
	WordMBSIBackgroundPreset10 = '\000\305\000\012',
	WordMBSIBackgroundPreset11 = '\000\305\000\013',
	WordMBSIBackgroundPreset12 = '\000\305\000\014'
};
typedef enum WordMBSI WordMBSI;

enum WordPDrT {
	WordPDrTTextDirectionUnset = '\000\352\377\376',
	WordPDrTLeftToRight = '\000\353\000\001',
	WordPDrTRightToLeft = '\000\353\000\002'
};
typedef enum WordPDrT WordPDrT;

enum WordPBtT {
	WordPBtTBulletTypeUnset = '\000\347\377\376',
	WordPBtTBulletTypeNone = '\000\350\000\000',
	WordPBtTBulletTypeUnnumbered = '\000\350\000\001',
	WordPBtTBulletTypeNumbered = '\000\350\000\002',
	WordPBtTPictureBulletType = '\000\350\000\003'
};
typedef enum WordPBtT WordPBtT;

enum WordPBtS {
	WordPBtSNumberedBulletStyleUnset = '\000\350\377\376',
	WordPBtSNumberedBulletStyleAlphaLowercasePeriod = '\000\351\000\000',
	WordPBtSNumberedBulletStyleAlphaUppercasePeriod = '\000\351\000\001',
	WordPBtSNumberedBulletStyleArabicRightParen = '\000\351\000\002',
	WordPBtSNumberedBulletStyleArabicPeriod = '\000\351\000\003',
	WordPBtSNumberedBulletStyleRomanLowercaseParenBoth = '\000\351\000\004',
	WordPBtSNumberedBulletStyleRomanLowercaseParenRight = '\000\351\000\005',
	WordPBtSNumberedBulletStyleRomanLowercasePeriod = '\000\351\000\006',
	WordPBtSNumberedBulletStyleRomanUppercasePeriod = '\000\351\000\007',
	WordPBtSNumberedBulletStyleAlphaLowercaseParenBoth = '\000\351\000\010',
	WordPBtSNumberedBulletStyleAlphaLowercaseParenRight = '\000\351\000\011',
	WordPBtSNumberedBulletStyleAlphaUppercaseParenBoth = '\000\351\000\012',
	WordPBtSNumberedBulletStyleAlphaUppercaseParenRight = '\000\351\000\013',
	WordPBtSNumberedBulletStyleArabicParenBoth = '\000\351\000\014',
	WordPBtSNumberedBulletStyleArabicPlain = '\000\351\000\015',
	WordPBtSNumberedBulletStyleRomanUppercaseParenBoth = '\000\351\000\016',
	WordPBtSNumberedBulletStyleRomanUppercaseParenRight = '\000\351\000\017',
	WordPBtSNumberedBulletStyleSimplifiedChinesePlain = '\000\351\000\020',
	WordPBtSNumberedBulletStyleSimplifiedChinesePeriod = '\000\351\000\021',
	WordPBtSNumberedBulletStyleCircleNumberPlain = '\000\351\000\022',
	WordPBtSNumberedBulletStyleCircleNumberWhitePlain = '\000\351\000\023',
	WordPBtSNumberedBulletStyleCircleNumberBlackPlain = '\000\351\000\024',
	WordPBtSNumberedBulletStyleTraditionalChinesePlain = '\000\351\000\025',
	WordPBtSNumberedBulletStyleTraditionalChinesePeriod = '\000\351\000\026',
	WordPBtSNumberedBulletStyleArabicAlphaDash = '\000\351\000\027',
	WordPBtSNumberedBulletStyleArabicAbjadDash = '\000\351\000\030',
	WordPBtSNumberedBulletStyleHebrewAlphaDash = '\000\351\000\031',
	WordPBtSNumberedBulletStyleKanjiKoreanPlain = '\000\351\000\032',
	WordPBtSNumberedBulletStyleKanjiKoreanPeriod = '\000\351\000\033',
	WordPBtSNumberedBulletStyleArabicDBPlain = '\000\351\000\034',
	WordPBtSNumberedBulletStyleArabicDBPeriod = '\000\351\000\035',
	WordPBtSNumberedBulletStyleThaiAlphaPeriod = '\000\351\000\036',
	WordPBtSNumberedBulletStyleThaiAlphaParenRight = '\000\351\000\037',
	WordPBtSNumberedBulletStyleThaiAlphaParenBoth = '\000\351\000 ',
	WordPBtSNumberedBulletStyleThaiNumberPeriod = '\000\351\000!',
	WordPBtSNumberedBulletStyleThaiNumberParenRight = '\000\351\000\"',
	WordPBtSNumberedBulletStyleThaiParenBoth = '\000\351\000#',
	WordPBtSNumberedBulletStyleHindiAlphaPeriod = '\000\351\000$',
	WordPBtSNumberedBulletStyleHindiNumberPeriod = '\000\351\000%',
	WordPBtSNumberedBulletStyleKanjiSimpifiedChineseDBPeriod = '\000\351\000&',
	WordPBtSNumberedBulletStyleHindiNumberParenRight = '\000\351\000\'',
	WordPBtSNumberedBulletStyleHindiAlpha1Period = '\000\351\000('
};
typedef enum WordPBtS WordPBtS;

enum WordPTSp {
	WordPTSpTabstopUnset = '\000\364\377\376',
	WordPTSpTabstopLeft = '\000\365\000\001',
	WordPTSpTabstopCenter = '\000\365\000\002',
	WordPTSpTabstopRight = '\000\365\000\003',
	WordPTSpTabstopDecimal = '\000\365\000\004'
};
typedef enum WordPTSp WordPTSp;

enum WordMRfT {
	WordMRfTReflectionUnset = '\003\351\377\376',
	WordMRfTReflectionTypeNone = '\003\352\000\000',
	WordMRfTReflectionType1 = '\003\352\000\001',
	WordMRfTReflectionType2 = '\003\352\000\002',
	WordMRfTReflectionType3 = '\003\352\000\003',
	WordMRfTReflectionType4 = '\003\352\000\004',
	WordMRfTReflectionType5 = '\003\352\000\005',
	WordMRfTReflectionType6 = '\003\352\000\006',
	WordMRfTReflectionType7 = '\003\352\000\007',
	WordMRfTReflectionType8 = '\003\352\000\010',
	WordMRfTReflectionType9 = '\003\352\000\011'
};
typedef enum WordMRfT WordMRfT;

enum WordMTtA {
	WordMTtATextureUnset = '\003\352\377\376',
	WordMTtATextureTopLeft = '\003\353\000\000',
	WordMTtATextureTop = '\003\353\000\001',
	WordMTtATextureTopRight = '\003\353\000\002',
	WordMTtATextureLeft = '\003\353\000\003',
	WordMTtATextureCenter = '\003\353\000\004',
	WordMTtATextureRight = '\003\353\000\005',
	WordMTtATextureBottomLeft = '\003\353\000\006',
	WordMTtATextureBotton = '\003\353\000\007',
	WordMTtATextureBottomRight = '\003\353\000\010'
};
typedef enum WordMTtA WordMTtA;

enum WordPBlA {
	WordPBlATextBaselineAlignmentUnset = '\003\353\377\376',
	WordPBlATextBaselineAlignBaseline = '\003\354\000\001',
	WordPBlATextBaselineAlignTop = '\003\354\000\002',
	WordPBlATextBaselineAlignCenter = '\003\354\000\003',
	WordPBlATextBaselineAlignEastAsian50 = '\003\354\000\004',
	WordPBlATextBaselineAlignAutomatic = '\003\354\000\005'
};
typedef enum WordPBlA WordPBlA;

enum WordMCbF {
	WordMCbFClipboardFormatUnset = '\003\354\377\376',
	WordMCbFNativeClipboardFormat = '\003\355\000\001',
	WordMCbFHTMlClipboardFormat = '\003\355\000\002',
	WordMCbFRTFClipboardFormat = '\003\355\000\003',
	WordMCbFPlainTextClipboardFormat = '\003\355\000\004'
};
typedef enum WordMCbF WordMCbF;

enum WordMTiP {
	WordMTiPInsertBefore = '\003\356\000\000',
	WordMTiPInsertAfter = '\003\356\000\001'
};
typedef enum WordMTiP WordMTiP;

enum WordMPiT {
	WordMPiTSaveAsDefault = '\003\362\377\376',
	WordMPiTSaveAsPNGFile = '\003\363\000\000',
	WordMPiTSaveAsBMPFile = '\003\363\000\001',
	WordMPiTSaveAsGIFFile = '\003\363\000\002',
	WordMPiTSaveAsJPGFile = '\003\363\000\003',
	WordMPiTSaveAsPDFFile = '\003\363\000\004'
};
typedef enum WordMPiT WordMPiT;

enum WordMPeT {
	WordMPeTNoEffect = '\003\364\000\000',
	WordMPeTEffectBackgroundRemoval = '\003\364\000\001',
	WordMPeTEffectBlur = '\003\364\000\002',
	WordMPeTEffectBrightnessContrast = '\003\364\000\003',
	WordMPeTEffectCement = '\003\364\000\004',
	WordMPeTEffectCrisscrossEtching = '\003\364\000\005',
	WordMPeTEffectChalkSketch = '\003\364\000\006',
	WordMPeTEffectColorTemperature = '\003\364\000\007',
	WordMPeTEffectCutout = '\003\364\000\010',
	WordMPeTEffectFilmGrain = '\003\364\000\011',
	WordMPeTEffectGlass = '\003\364\000\012',
	WordMPeTEffectGlowDiffused = '\003\364\000\013',
	WordMPeTEffectGlowEdges = '\003\364\000\014',
	WordMPeTEffectLightScreen = '\003\364\000\015',
	WordMPeTEffectLineDrawing = '\003\364\000\016',
	WordMPeTEffectMarker = '\003\364\000\017',
	WordMPeTEffectMosiaicBubbles = '\003\364\000\020',
	WordMPeTEffectPaintBrush = '\003\364\000\021',
	WordMPeTEffectPaintStrokes = '\003\364\000\022',
	WordMPeTEffectPastelsSmooth = '\003\364\000\023',
	WordMPeTEffectPencilGrayscale = '\003\364\000\024',
	WordMPeTEffectPencilSketch = '\003\364\000\025',
	WordMPeTEffectPhotocopy = '\003\364\000\026',
	WordMPeTEffectPlasticWrap = '\003\364\000\027',
	WordMPeTEffectSaturation = '\003\364\000\030',
	WordMPeTEffectSharpenSoften = '\003\364\000\031',
	WordMPeTEffectTexturizer = '\003\364\000\032',
	WordMPeTEffectWatercolorSponge = '\003\364\000\033'
};
typedef enum WordMPeT WordMPeT;

enum WordMSgT {
	WordMSgTLine = '\000\217\000\000',
	WordMSgTCurve = '\000\217\000\001'
};
typedef enum WordMSgT WordMSgT;

enum WordMEdT {
	WordMEdTAuto = '\000\220\000\000',
	WordMEdTCorner = '\000\220\000\001',
	WordMEdTSmooth = '\000\220\000\002',
	WordMEdTSymmetric = '\000\220\000\003'
};
typedef enum WordMEdT WordMEdT;

enum WordSANP {
	WordSANPDefaultNodePosition = '\003\365\000\001',
	WordSANPAfterNode = '\003\365\000\002',
	WordSANPBeforeNode = '\003\365\000\003',
	WordSANPAboveNode = '\003\365\000\004',
	WordSANPBelowNode = '\003\365\000\005'
};
typedef enum WordSANP WordSANP;

enum WordSANT {
	WordSANTDefaultNode = '\003\366\000\001',
	WordSANTAssistantNode = '\003\366\000\002'
};
typedef enum WordSANT WordSANT;

enum WordMOCT {
	WordMOCTOrgChartLayoutUnset = '\003\366\377\376',
	WordMOCTOrgChartLayoutStandard = '\003\367\000\001',
	WordMOCTOrgChartLayoutBothHanging = '\003\367\000\002',
	WordMOCTOrgChartLayoutLeftHanging = '\003\367\000\003',
	WordMOCTOrgChartLayoutRightHanging = '\003\367\000\004',
	WordMOCTOrgChartLayoutDefault = '\003\367\000\005'
};
typedef enum WordMOCT WordMOCT;

enum WordMAlC {
	WordMAlCAlignLefts = '\000\000\000\000',
	WordMAlCAlignCenters = '\000\000\000\001',
	WordMAlCAlignRights = '\000\000\000\002',
	WordMAlCAlignTops = '\000\000\000\003',
	WordMAlCAlignMiddles = '\000\000\000\004',
	WordMAlCAlignBottoms = '\000\000\000\005'
};
typedef enum WordMAlC WordMAlC;

enum WordMDsC {
	WordMDsCDistributeHorizontally = '\000\000\000\000',
	WordMDsCDistributeVertically = '\000\000\000\001'
};
typedef enum WordMDsC WordMDsC;

enum WordMOrT {
	WordMOrTOrientationUnset = '\000\214\377\376',
	WordMOrTHorizontalOrientation = '\000\215\000\001',
	WordMOrTVerticalOrientation = '\000\215\000\002'
};
typedef enum WordMOrT WordMOrT;

enum WordMZoC {
	WordMZoCBringShapeToFront = '\000p\000\000',
	WordMZoCSendShapeToBack = '\000p\000\001',
	WordMZoCBringShapeForward = '\000p\000\002',
	WordMZoCSendShapeBackward = '\000p\000\003',
	WordMZoCBringShapeInFrontOfText = '\000p\000\004',
	WordMZoCSendShapeBehindText = '\000p\000\005'
};
typedef enum WordMZoC WordMZoC;

enum WordMFlC {
	WordMFlCFlipHorizontal = '\000q\000\000',
	WordMFlCFlipVertical = '\000q\000\001'
};
typedef enum WordMFlC WordMFlC;

enum WordMTri {
	WordMTriTrue = '\000\240\377\377',
	WordMTriFalse = '\000\241\000\000',
	WordMTriCTrue = '\000\241\000\001',
	WordMTriToggle = '\000\240\377\375',
	WordMTriTriStateUnset = '\000\240\377\376'
};
typedef enum WordMTri WordMTri;

enum WordMBW {
	WordMBWBlackAndWhiteUnset = '\000\254\377\376',
	WordMBWBlackAndWhiteModeAutomatic = '\000\255\000\001',
	WordMBWBlackAndWhiteModeGrayScale = '\000\255\000\002',
	WordMBWBlackAndWhiteModeLightGrayScale = '\000\255\000\003',
	WordMBWBlackAndWhiteModeInverseGrayScale = '\000\255\000\004',
	WordMBWBlackAndWhiteModeGrayOutline = '\000\255\000\005',
	WordMBWBlackAndWhiteModeBlackTextAndLine = '\000\255\000\006',
	WordMBWBlackAndWhiteModeHighContrast = '\000\255\000\007',
	WordMBWBlackAndWhiteModeBlack = '\000\255\000\010',
	WordMBWBlackAndWhiteModeWhite = '\000\255\000\011',
	WordMBWBlackAndWhiteModeDontShow = '\000\255\000\012'
};
typedef enum WordMBW WordMBW;

enum WordMBPS {
	WordMBPSBarLeft = '\000r\000\000',
	WordMBPSBarTop = '\000r\000\001',
	WordMBPSBarRight = '\000r\000\002',
	WordMBPSBarBottom = '\000r\000\003',
	WordMBPSBarFloating = '\000r\000\004',
	WordMBPSBarPopUp = '\000r\000\005',
	WordMBPSBarMenu = '\000r\000\006'
};
typedef enum WordMBPS WordMBPS;

enum WordMBPt {
	WordMBPtNoProtection = '\000s\000\000',
	WordMBPtNoCustomize = '\000s\000\001',
	WordMBPtNoResize = '\000s\000\002',
	WordMBPtNoMove = '\000s\000\004',
	WordMBPtNoChangeVisible = '\000s\000\010',
	WordMBPtNoChangeDock = '\000s\000\020',
	WordMBPtNoVerticalDock = '\000s\000 ',
	WordMBPtNoHorizontalDock = '\000s\000@'
};
typedef enum WordMBPt WordMBPt;

enum WordMBTY {
	WordMBTYNormalCommandBar = '\000t\000\000',
	WordMBTYMenubarCommandBar = '\000t\000\001',
	WordMBTYPopupCommandBar = '\000t\000\002'
};
typedef enum WordMBTY WordMBTY;

enum WordMCLT {
	WordMCLTControlCustom = '\000u\000\000',
	WordMCLTControlButton = '\000u\000\001',
	WordMCLTControlEdit = '\000u\000\002',
	WordMCLTControlDropDown = '\000u\000\003',
	WordMCLTControlCombobox = '\000u\000\004',
	WordMCLTButtonDropDown = '\000u\000\005',
	WordMCLTSplitDropDown = '\000u\000\006',
	WordMCLTOCXDropDown = '\000u\000\007',
	WordMCLTGenericDropDown = '\000u\000\010',
	WordMCLTGraphicDropDown = '\000u\000\011',
	WordMCLTControlPopup = '\000u\000\012',
	WordMCLTGraphicPopup = '\000u\000\013',
	WordMCLTButtonPopup = '\000u\000\014',
	WordMCLTSplitButtonPopup = '\000u\000\015',
	WordMCLTSplitButtonMRUPopup = '\000u\000\016',
	WordMCLTControlLabel = '\000u\000\017',
	WordMCLTExpandingGrid = '\000u\000\020',
	WordMCLTSplitExpandingGrid = '\000u\000\021',
	WordMCLTControlGrid = '\000u\000\022',
	WordMCLTControlGauge = '\000u\000\023',
	WordMCLTGraphicCombobox = '\000u\000\024',
	WordMCLTControlPane = '\000u\000\025',
	WordMCLTActiveX = '\000u\000\026',
	WordMCLTControlGroup = '\000u\000\027',
	WordMCLTControlTab = '\000u\000\030',
	WordMCLTControlSpinner = '\000u\000\031'
};
typedef enum WordMCLT WordMCLT;

enum WordMBns {
	WordMBnsButtonStateUp = '\000v\000\000',
	WordMBnsButtonStateDown = '\000u\377\377',
	WordMBnsButtonStateUnset = '\000v\000\002'
};
typedef enum WordMBns WordMBns;

enum WordMcOu {
	WordMcOuNeither = '\000w\000\000',
	WordMcOuServer = '\000w\000\001',
	WordMcOuClient = '\000w\000\002',
	WordMcOuBoth = '\000w\000\003'
};
typedef enum WordMcOu WordMcOu;

enum WordMBTs {
	WordMBTsButtonAutomatic = '\000x\000\000',
	WordMBTsButtonIcon = '\000x\000\001',
	WordMBTsButtonCaption = '\000x\000\002',
	WordMBTsButtonIconAndCaption = '\000x\000\003'
};
typedef enum WordMBTs WordMBTs;

enum WordMXcb {
	WordMXcbComboboxStyleNormal = '\000y\000\000',
	WordMXcbComboboxStyleLabel = '\000y\000\001'
};
typedef enum WordMXcb WordMXcb;

enum WordMMNA {
	WordMMNANone = '\000{\000\000',
	WordMMNARandom = '\000{\000\001',
	WordMMNAUnfold = '\000{\000\002',
	WordMMNASlide = '\000{\000\003'
};
typedef enum WordMMNA WordMMNA;

enum WordMHlT {
	WordMHlTHyperlinkTypeTextRange = '\000\226\000\000',
	WordMHlTHyperlinkTypeShape = '\000\226\000\001',
	WordMHlTHyperlinkTypeInlineShape = '\000\226\000\002'
};
typedef enum WordMHlT WordMHlT;

enum WordMXiM {
	WordMXiMAppendString = '\000\256\000\000',
	WordMXiMPostString = '\000\256\000\001'
};
typedef enum WordMXiM WordMXiM;

enum WordMANT {
	WordMANTIdle = '\000|\000\001',
	WordMANTGreeting = '\000|\000\002',
	WordMANTGoodbye = '\000|\000\003',
	WordMANTBeginSpeaking = '\000|\000\004',
	WordMANTCharacterSuccessMajor = '\000|\000\006',
	WordMANTGetAttentionMajor = '\000|\000\013',
	WordMANTGetAttentionMinor = '\000|\000\014',
	WordMANTSearching = '\000|\000\015',
	WordMANTPrinting = '\000|\000\022',
	WordMANTGestureRight = '\000|\000\023',
	WordMANTWritingNotingSomething = '\000|\000\026',
	WordMANTWorkingAtSomething = '\000|\000\027',
	WordMANTThinking = '\000|\000\030',
	WordMANTSendingMail = '\000|\000\031',
	WordMANTListensToComputer = '\000|\000\032',
	WordMANTDisappear = '\000|\000\037',
	WordMANTAppear = '\000|\000 ',
	WordMANTGetArtsy = '\000|\000d',
	WordMANTGetTechy = '\000|\000e',
	WordMANTGetWizardy = '\000|\000f',
	WordMANTCheckingSomething = '\000|\000g',
	WordMANTLookDown = '\000|\000h',
	WordMANTLookDownLeft = '\000|\000i',
	WordMANTLookDownRight = '\000|\000j',
	WordMANTLookLeft = '\000|\000k',
	WordMANTLookRight = '\000|\000l',
	WordMANTLookUp = '\000|\000m',
	WordMANTLookUpLeft = '\000|\000n',
	WordMANTLookUpRight = '\000|\000o',
	WordMANTSaving = '\000|\000p',
	WordMANTGestureDown = '\000|\000q',
	WordMANTGestureLeft = '\000|\000r',
	WordMANTGestureUp = '\000|\000s',
	WordMANTEmptyTrash = '\000|\000t'
};
typedef enum WordMANT WordMANT;

enum WordMBSt {
	WordMBStButtonNone = '\000}\000\000',
	WordMBStButtonOk = '\000}\000\001',
	WordMBStButtonCancel = '\000}\000\002',
	WordMBStButtonsOkCancel = '\000}\000\003',
	WordMBStButtonsYesNo = '\000}\000\004',
	WordMBStButtonsYesNoCancel = '\000}\000\005',
	WordMBStButtonsBackClose = '\000}\000\006',
	WordMBStButtonsNextClose = '\000}\000\007',
	WordMBStButtonsBackNextClose = '\000}\000\010',
	WordMBStButtonsRetryCancel = '\000}\000\011',
	WordMBStButtonsAbortRetryIgnore = '\000}\000\012',
	WordMBStButtonsSearchClose = '\000}\000\013',
	WordMBStButtonsBackNextSnooze = '\000}\000\014',
	WordMBStButtonsTipsOptionsClose = '\000}\000\015',
	WordMBStButtonsYesAllNoCancel = '\000}\000\016'
};
typedef enum WordMBSt WordMBSt;

enum WordMIct {
	WordMIctIconNone = '\000~\000\000',
	WordMIctIconApplication = '\000~\000\001',
	WordMIctIconAlert = '\000~\000\002',
	WordMIctIconTip = '\000~\000\003',
	WordMIctIconAlertCritical = '\000~\000e',
	WordMIctIconAlertWarning = '\000~\000g',
	WordMIctIconAlertInfo = '\000~\000h'
};
typedef enum WordMIct WordMIct;

enum WordMWAt {
	WordMWAtInactive = '\000\202\000\000',
	WordMWAtActive = '\000\202\000\001',
	WordMWAtSuspend = '\000\202\000\002',
	WordMWAtResume = '\000\202\000\003'
};
typedef enum WordMWAt WordMWAt;

enum WordMeDP {
	WordMeDPPropertyTypeNumber = '\000\242\000\001',
	WordMeDPPropertyTypeBoolean = '\000\242\000\002',
	WordMeDPPropertyTypeDate = '\000\242\000\003',
	WordMeDPPropertyTypeString = '\000\242\000\004',
	WordMeDPPropertyTypeFloat = '\000\242\000\005'
};
typedef enum WordMeDP WordMeDP;

enum WordMASc {
	WordMAScMsoAutomationSecurityLow = '\000\243\000\001',
	WordMAScMsoAutomationSecurityByUI = '\000\243\000\002',
	WordMAScMsoAutomationSecurityForceDisable = '\000\243\000\003'
};
typedef enum WordMASc WordMASc;

enum WordMSsz {
	WordMSszResolution544x376 = '\000\204\000\000',
	WordMSszResolution640x480 = '\000\204\000\001',
	WordMSszResolution720x512 = '\000\204\000\002',
	WordMSszResolution800x600 = '\000\204\000\003',
	WordMSszResolution1024x768 = '\000\204\000\004',
	WordMSszResolution1152x882 = '\000\204\000\005',
	WordMSszResolution1152x900 = '\000\204\000\006',
	WordMSszResolution1280x1024 = '\000\204\000\007',
	WordMSszResolution1600x1200 = '\000\204\000\010',
	WordMSszResolution1800x1440 = '\000\204\000\011',
	WordMSszResolution1920x1200 = '\000\204\000\012'
};
typedef enum WordMSsz WordMSsz;

enum WordMChS {
	WordMChSArabicCharacterSet = '\000\205\000\001',
	WordMChSCyrillicCharacterSet = '\000\205\000\002',
	WordMChSEnglishCharacterSet = '\000\205\000\003',
	WordMChSGreekCharacterSet = '\000\205\000\004',
	WordMChSHebrewCharacterSet = '\000\205\000\005',
	WordMChSJapaneseCharacterSet = '\000\205\000\006',
	WordMChSKoreanCharacterSet = '\000\205\000\007',
	WordMChSMultilingualUnicodeCharacterSet = '\000\205\000\010',
	WordMChSSimplifiedChineseCharacterSet = '\000\205\000\011',
	WordMChSThaiCharacterSet = '\000\205\000\012',
	WordMChSTraditionalChineseCharacterSet = '\000\205\000\013',
	WordMChSVietnameseCharacterSet = '\000\205\000\014'
};
typedef enum WordMChS WordMChS;

enum WordMtEn {
	WordMtEnEncodingThai = '\000\213\003j',
	WordMtEnEncodingJapaneseShiftJIS = '\000\213\003\244',
	WordMtEnEncodingSimplifiedChinese = '\000\213\003\250',
	WordMtEnEncodingKorean = '\000\213\003\265',
	WordMtEnEncodingBig5TraditionalChinese = '\000\213\003\266',
	WordMtEnEncodingLittleEndian = '\000\213\004\260',
	WordMtEnEncodingBigEndian = '\000\213\004\261',
	WordMtEnEncodingCentralEuropean = '\000\213\004\342',
	WordMtEnEncodingCyrillic = '\000\213\004\343',
	WordMtEnEncodingWestern = '\000\213\004\344',
	WordMtEnEncodingGreek = '\000\213\004\345',
	WordMtEnEncodingTurkish = '\000\213\004\346',
	WordMtEnEncodingHebrew = '\000\213\004\347',
	WordMtEnEncodingArabic = '\000\213\004\350',
	WordMtEnEncodingBaltic = '\000\213\004\351',
	WordMtEnEncodingVietnamese = '\000\213\004\352',
	WordMtEnEncodingISO88591Latin1 = '\000\213o\257',
	WordMtEnEncodingISO88592CentralEurope = '\000\213o\260',
	WordMtEnEncodingISO88593Latin3 = '\000\213o\261',
	WordMtEnEncodingISO88594Baltic = '\000\213o\262',
	WordMtEnEncodingISO88595Cyrillic = '\000\213o\263',
	WordMtEnEncodingISO88596Arabic = '\000\213o\264',
	WordMtEnEncodingISO88597Greek = '\000\213o\265',
	WordMtEnEncodingISO88598Hebrew = '\000\213o\266',
	WordMtEnEncodingISO88599Turkish = '\000\213o\267',
	WordMtEnEncodingISO885915Latin9 = '\000\213o\275',
	WordMtEnEncodingISO2022JapaneseNoHalfWidthKatakana = '\000\213\304,',
	WordMtEnEncodingISO2022JapaneseJISX02021984 = '\000\213\304-',
	WordMtEnEncodingISO2022JapaneseJISX02011989 = '\000\213\304.',
	WordMtEnEncodingISO2022KR = '\000\213\3041',
	WordMtEnEncodingISO2022CNTraditionalChinese = '\000\213\3043',
	WordMtEnEncodingISO2022CNSimplifiedChinese = '\000\213\3045',
	WordMtEnEncodingMacRoman = '\000\213\'\020',
	WordMtEnEncodingMacJapanese = '\000\213\'\021',
	WordMtEnEncodingMacTraditionalChinese = '\000\213\'\022',
	WordMtEnEncodingMacKorean = '\000\213\'\023',
	WordMtEnEncodingMacArabic = '\000\213\'\024',
	WordMtEnEncodingMacHebrew = '\000\213\'\025',
	WordMtEnEncodingMacGreek1 = '\000\213\'\026',
	WordMtEnEncodingMacCyrillic = '\000\213\'\027',
	WordMtEnEncodingMacSimplifiedChineseGB2312 = '\000\213\'\030',
	WordMtEnEncodingMacRomania = '\000\213\'\032',
	WordMtEnEncodingMacUkraine = '\000\213\'!',
	WordMtEnEncodingMacLatin2 = '\000\213\'-',
	WordMtEnEncodingMacIcelandic = '\000\213\'_',
	WordMtEnEncodingMacTurkish = '\000\213\'a',
	WordMtEnEncodingMacCroatia = '\000\213\'b',
	WordMtEnEncodingEBCDICUSCanada = '\000\213\000%',
	WordMtEnEncodingEBCDICInternational = '\000\213\001\364',
	WordMtEnEncodingEBCDICMultilingualROECELatin2 = '\000\213\003f',
	WordMtEnEncodingEBCDICGreekModern = '\000\213\003k',
	WordMtEnEncodingEBCDICTurkishLatin5 = '\000\213\004\002',
	WordMtEnEncodingEBCDICGermany = '\000\213O1',
	WordMtEnEncodingEBCDICDenmarkNorway = '\000\213O5',
	WordMtEnEncodingEBCDICFinlandSweden = '\000\213O6',
	WordMtEnEncodingEBCDICItaly = '\000\213O8',
	WordMtEnEncodingEBCDICLatinAmericaSpain = '\000\213O<',
	WordMtEnEncodingEBCDICUnitedKingdom = '\000\213O=',
	WordMtEnEncodingEBCDICJapaneseKatakanaExtended = '\000\213OB',
	WordMtEnEncodingEBCDICFrance = '\000\213OI',
	WordMtEnEncodingEBCDICArabic = '\000\213O\304',
	WordMtEnEncodingEBCDICGreek = '\000\213O\307',
	WordMtEnEncodingEBCDICHebrew = '\000\213O\310',
	WordMtEnEncodingEBCDICKoreanExtended = '\000\213Qa',
	WordMtEnEncodingEBCDICThai = '\000\213Qf',
	WordMtEnEncodingEBCDICIcelandic = '\000\213Q\207',
	WordMtEnEncodingEBCDICTurkish = '\000\213Q\251',
	WordMtEnEncodingEBCDICRussian = '\000\213Q\220',
	WordMtEnEncodingEBCDICSerbianBulgarian = '\000\213R!',
	WordMtEnEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = '\000\213\306\362',
	WordMtEnEncodingEBCDICUSCanadaAndJapanese = '\000\213\306\363',
	WordMtEnEncodingEBCDICExtendedAndKorean = '\000\213\306\365',
	WordMtEnEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = '\000\213\306\367',
	WordMtEnEncodingEBCDICUSCanadaAndTraditionalChinese = '\000\213\306\371',
	WordMtEnEncodingEBCDICJapaneseLatinExtendedAndJapanese = '\000\213\306\373',
	WordMtEnEncodingOEMUnitedStates = '\000\213\001\265',
	WordMtEnEncodingOEMGreek = '\000\213\002\341',
	WordMtEnEncodingOEMBaltic = '\000\213\003\007',
	WordMtEnEncodingOEMMultilingualLatinI = '\000\213\003R',
	WordMtEnEncodingOEMMultilingualLatinII = '\000\213\003T',
	WordMtEnEncodingOEMCyrillic = '\000\213\003W',
	WordMtEnEncodingOEMTurkish = '\000\213\003Y',
	WordMtEnEncodingOEMPortuguese = '\000\213\003\\',
	WordMtEnEncodingOEMIcelandic = '\000\213\003]',
	WordMtEnEncodingOEMHebrew = '\000\213\003^',
	WordMtEnEncodingOEMCanadianFrench = '\000\213\003_',
	WordMtEnEncodingOEMArabic = '\000\213\003`',
	WordMtEnEncodingOEMNordic = '\000\213\003a',
	WordMtEnEncodingOEMCyrillicII = '\000\213\003b',
	WordMtEnEncodingOEMModernGreek = '\000\213\003e',
	WordMtEnEncodingEUCJapanese = '\000\213\312\334',
	WordMtEnEncodingEUCChineseSimplifiedChinese = '\000\213\312\340',
	WordMtEnEncodingEUCKorean = '\000\213\312\355',
	WordMtEnEncodingEUCTaiwaneseTraditionalChinese = '\000\213\312\356',
	WordMtEnEncodingDevanagari = '\000\213\336\252',
	WordMtEnEncodingBengali = '\000\213\336\253',
	WordMtEnEncodingTamil = '\000\213\336\254',
	WordMtEnEncodingTelugu = '\000\213\336\255',
	WordMtEnEncodingAssamese = '\000\213\336\256',
	WordMtEnEncodingOriya = '\000\213\336\257',
	WordMtEnEncodingKannada = '\000\213\336\260',
	WordMtEnEncodingMalayalam = '\000\213\336\261',
	WordMtEnEncodingGujarati = '\000\213\336\262',
	WordMtEnEncodingPunjabi = '\000\213\336\263',
	WordMtEnEncodingArabicASMO = '\000\213\002\304',
	WordMtEnEncodingArabicTransparentASMO = '\000\213\002\320',
	WordMtEnEncodingKoreanJohab = '\000\213\005Q',
	WordMtEnEncodingTaiwanCNS = '\000\213N ',
	WordMtEnEncodingTaiwanTCA = '\000\213N!',
	WordMtEnEncodingTaiwanEten = '\000\213N\"',
	WordMtEnEncodingTaiwanIBM5550 = '\000\213N#',
	WordMtEnEncodingTaiwanTeletext = '\000\213N$',
	WordMtEnEncodingTaiwanWang = '\000\213N%',
	WordMtEnEncodingIA5IRV = '\000\213N\211',
	WordMtEnEncodingIA5German = '\000\213N\212',
	WordMtEnEncodingIA5Swedish = '\000\213N\213',
	WordMtEnEncodingIA5Norwegian = '\000\213N\214',
	WordMtEnEncodingUSASCII = '\000\213N\237',
	WordMtEnEncodingT61 = '\000\213O%',
	WordMtEnEncodingISO6937NonspacingAccent = '\000\213O-',
	WordMtEnEncodingKOI8R = '\000\213Q\202',
	WordMtEnEncodingExtAlphaLowercase = '\000\213R#',
	WordMtEnEncodingKOI8U = '\000\213Uj',
	WordMtEnEncodingEuropa3 = '\000\213qI',
	WordMtEnEncodingHZGBSimplifiedChinese = '\000\213\316\310',
	WordMtEnEncodingUTF7 = '\000\213\375\350',
	WordMtEnEncodingUTF8 = '\000\213\375\351'
};
typedef enum WordMtEn WordMtEn;

enum Word4000 {
	Word4000CommandBar = 'msCB',
	Word4000CommandBarControl = 'mCBC'
};
typedef enum Word4000 Word4000;

enum WordEMvS {
	WordEMvSTowardsTheEnd = '\002\266\000\000',
	WordEMvSTowardsTheStart = '\002\266\000\001',
	WordEMvSTowardsTheLeft = '\002\266\000\002',
	WordEMvSTowardsTheRight = '\002\266\000\003',
	WordEMvSTowardsTheTop = '\002\266\000\004',
	WordEMvSTowardsTheBottom = '\002\266\000\005'
};
typedef enum WordEMvS WordEMvS;

enum WordEMUW {
	WordEMUWTowardsTheTail = '\002\267\000\000',
	WordEMUWTowardsTheBeginning = '\002\267\000\001'
};
typedef enum WordEMUW WordEMUW;

enum WordEiAb {
	WordEiAbAbove = '\002\270\000\000',
	WordEiAbBelow = '\002\270\000\001'
};
typedef enum WordEiAb WordEiAb;

enum WordEiBa {
	WordEiBaBeforeTheObject = '\002\271\000\000',
	WordEiBaAfterTheObject = '\002\271\000\001'
};
typedef enum WordEiBa WordEiBa;

enum WordEiRL {
	WordEiRLInsertOnTheRight = '\002\273\000\000',
	WordEiRLInsertOnTheLeft = '\002\273\000\001'
};
typedef enum WordEiRL WordEiRL;

enum WordEIPt {
	WordEIPtOffset = '\003\037\000\000',
	WordEIPtLocationReference = '\003\037\000\001'
};
typedef enum WordEIPt WordEIPt;

enum WordEFRt {
	WordEFRtTextRange = '\003\036\000\000',
	WordEFRtInsertionPoint = '\003\036\000\001'
};
typedef enum WordEFRt WordEFRt;

enum WordE101 {
	WordE101TypeNormalTemplate = '\001\365\000\000',
	WordE101TypeGlobalTemplate = '\001\365\000\001',
	WordE101TypeAttachedTemplate = '\001\365\000\002'
};
typedef enum WordE101 WordE101;

enum WordE102 {
	WordE102ContinueDisabled = '\001\366\000\000',
	WordE102ResetList = '\001\366\000\001',
	WordE102ContinueList = '\001\366\000\002'
};
typedef enum WordE102 WordE102;

enum WordE103 {
	WordE103IMEModeNoControl = '\001\367\000\000',
	WordE103IMEModeOn = '\001\367\000\001',
	WordE103IMEModeOff = '\001\367\000\002',
	WordE103IMEModeHiragana = '\001\367\000\004',
	WordE103IMEModeKatakana = '\001\367\000\005',
	WordE103IMEModeKatakanaHalf = '\001\367\000\006',
	WordE103IMEModeAlphaFull = '\001\367\000\007',
	WordE103IMEModeAlpha = '\001\367\000\010',
	WordE103IMEModeHangulFull = '\001\367\000\011',
	WordE103IMEModeHangul = '\001\367\000\012'
};
typedef enum WordE103 WordE103;

enum WordE104 {
	WordE104BaselineAlignTop = '\001\370\000\000',
	WordE104BaselineAlignCenter = '\001\370\000\001',
	WordE104BaselineAlignBaseline = '\001\370\000\002',
	WordE104BaselineAlignEastAsian50 = '\001\370\000\003',
	WordE104BaselineAlignAuto = '\001\370\000\004'
};
typedef enum WordE104 WordE104;

enum WordE105 {
	WordE105IndexFilterNone = '\001\371\000\000',
	WordE105IndexFilterAiueo = '\001\371\000\001',
	WordE105IndexFilterAkasatana = '\001\371\000\002',
	WordE105IndexFilterChosung = '\001\371\000\003',
	WordE105IndexFilterLow = '\001\371\000\004',
	WordE105IndexFilterMedium = '\001\371\000\005',
	WordE105IndexFilterFull = '\001\371\000\006'
};
typedef enum WordE105 WordE105;

enum WordE106 {
	WordE106IndexSortByStroke = '\001\372\000\000',
	WordE106IndexSortBySyllable = '\001\372\000\001'
};
typedef enum WordE106 WordE106;

enum WordE107 {
	WordE107JustificationModeExpand = '\001\373\000\000',
	WordE107JustificationModeCompress = '\001\373\000\001',
	WordE107JustificationModeCompressKana = '\001\373\000\002'
};
typedef enum WordE107 WordE107;

enum WordE108 {
	WordE108EastAsianLineBreakLevelNormal = '\001\374\000\000',
	WordE108EastAsianLineBreakLevelStrict = '\001\374\000\001',
	WordE108EastAsianLineBreakLevelCustom = '\001\374\000\002'
};
typedef enum WordE108 WordE108;

enum WordE109 {
	WordE109HangulToHanja = '\001\375\000\000',
	WordE109HanjaToHangul = '\001\375\000\001'
};
typedef enum WordE109 WordE109;

enum WordE110 {
	WordE110Auto = '\001\376\000\000',
	WordE110Black = '\001\376\000\001',
	WordE110Blue = '\001\376\000\002',
	WordE110Turquoise = '\001\376\000\003',
	WordE110BrightGreen = '\001\376\000\004',
	WordE110Pink = '\001\376\000\005',
	WordE110Red = '\001\376\000\006',
	WordE110Yellow = '\001\376\000\007',
	WordE110White = '\001\376\000\010',
	WordE110DarkBlue = '\001\376\000\011',
	WordE110Teal = '\001\376\000\012',
	WordE110Green = '\001\376\000\013',
	WordE110Violet = '\001\376\000\014',
	WordE110DarkRed = '\001\376\000\015',
	WordE110DarkYellow = '\001\376\000\016',
	WordE110Gray50 = '\001\376\000\017',
	WordE110Gray25 = '\001\376\000\020',
	WordE110ByAuthor = '\001\375\377\377',
	WordE110NoHighlight = '\001\376\000\000'
};
typedef enum WordE110 WordE110;

enum WordE111 {
	WordE111ColorAutomatic = '\377\000\000\000',
	WordE111ColorBlack = '\000\000\000\000',
	WordE111ColorBlue = '\000\377\000\000',
	WordE111ColorPink = '\000\377\000\377',
	WordE111ColorRed = '\000\000\000\377',
	WordE111ColorYellow = '\000\000\377\377',
	WordE111ColorTurquoise = '\000\377\377\000',
	WordE111ColorBrightGreen = '\000\000\377\000',
	WordE111ColorGreen = '\000\000\200\000',
	WordE111ColorWhite = '\000\377\377\377',
	WordE111ColorDarkBlue = '\000\200\000\000',
	WordE111ColorTeal = '\000\200\200\000',
	WordE111ColorViolet = '\000\200\000\200',
	WordE111ColorDarkGreen = '\000\0003\000',
	WordE111ColorDarkRed = '\000\000\000\200',
	WordE111ColorDarkYellow = '\000\000\200\200',
	WordE111ColorBrown = '\000\0003\231',
	WordE111ColorOliveGreen = '\000\00033',
	WordE111ColorDarkTeal = '\000f3\000',
	WordE111ColorIndigo = '\000\23133',
	WordE111ColorOrange = '\000\000f\377',
	WordE111ColorBlueGray = '\000\231ff',
	WordE111ColorLightOrange = '\000\000\231\377',
	WordE111ColorLime = '\000\000\314\231',
	WordE111ColorSeaGreen = '\000f\2313',
	WordE111ColorAqua = '\000\314\3143',
	WordE111ColorLightBlue = '\000\377f3',
	WordE111ColorGold = '\000\000\314\377',
	WordE111ColorSkyBlue = '\000\377\314\000',
	WordE111ColorPlum = '\000f3\231',
	WordE111ColorRose = '\000\314\231\377',
	WordE111ColorTan = '\000\231\314\377',
	WordE111ColorLightYellow = '\000\231\377\377',
	WordE111ColorLightGreen = '\000\314\377\314',
	WordE111ColorLightTurquoise = '\000\377\377\314',
	WordE111ColorPaleBlue = '\000\377\314\231',
	WordE111ColorLavender = '\000\377\231\314',
	WordE111ColorGray05 = '\000\363\363\363',
	WordE111ColorGray10 = '\000\346\346\346',
	WordE111ColorGray125 = '\000\340\340\340',
	WordE111ColorGray15 = '\000\331\331\331',
	WordE111ColorGray20 = '\000\314\314\314',
	WordE111ColorGray25 = '\000\300\300\300',
	WordE111ColorGray30 = '\000\263\263\263',
	WordE111ColorGray35 = '\000\246\246\246',
	WordE111ColorGray375 = '\000\240\240\240',
	WordE111ColorGray40 = '\000\231\231\231',
	WordE111ColorGray45 = '\000\214\214\214',
	WordE111ColorGray50 = '\000\200\200\200',
	WordE111ColorGray55 = '\000sss',
	WordE111ColorGray60 = '\000fff',
	WordE111ColorGray625 = '\000```',
	WordE111ColorGray65 = '\000YYY',
	WordE111ColorGray70 = '\000LLL',
	WordE111ColorGray75 = '\000@@@',
	WordE111ColorGray80 = '\000333',
	WordE111ColorGray85 = '\000&&&',
	WordE111ColorGray875 = '\000   ',
	WordE111ColorGray90 = '\000\031\031\031',
	WordE111ColorGray95 = '\000\014\014\014'
};
typedef enum WordE111 WordE111;

enum WordE112 {
	WordE112TextureNone = '\002\000\000\000',
	WordE112Texture2Pt5Percent = '\002\000\000\031',
	WordE112Texture5Percent = '\002\000\0002',
	WordE112Texture7Pt5Percent = '\002\000\000K',
	WordE112Texture10Percent = '\002\000\000d',
	WordE112Texture12Pt5Percent = '\002\000\000}',
	WordE112Texture15Percent = '\002\000\000\226',
	WordE112Texture17Pt5Percent = '\002\000\000\257',
	WordE112Texture20Percent = '\002\000\000\310',
	WordE112Texture22Pt5Percent = '\002\000\000\341',
	WordE112Texture25Percent = '\002\000\000\372',
	WordE112Texture27Pt5Percent = '\002\000\001\023',
	WordE112Texture30Percent = '\002\000\001,',
	WordE112Texture32Pt5Percent = '\002\000\001E',
	WordE112Texture35Percent = '\002\000\001^',
	WordE112Texture37Pt5Percent = '\002\000\001w',
	WordE112Texture40Percent = '\002\000\001\220',
	WordE112Texture42Pt5Percent = '\002\000\001\251',
	WordE112Texture45Percent = '\002\000\001\302',
	WordE112Texture47Pt5Percent = '\002\000\001\333',
	WordE112Texture50Percent = '\002\000\001\364',
	WordE112Texture52Pt5Percent = '\002\000\002\015',
	WordE112Texture55Percent = '\002\000\002&',
	WordE112Texture57Pt5Percent = '\002\000\002\?',
	WordE112Texture60Percent = '\002\000\002X',
	WordE112Texture62Pt5Percent = '\002\000\002q',
	WordE112Texture65Percent = '\002\000\002\212',
	WordE112Texture67Pt5Percent = '\002\000\002\243',
	WordE112Texture70Percent = '\002\000\002\274',
	WordE112Texture72Pt5Percent = '\002\000\002\325',
	WordE112Texture75Percent = '\002\000\002\356',
	WordE112Texture77Pt5Percent = '\002\000\003\007',
	WordE112Texture80Percent = '\002\000\003 ',
	WordE112Texture82Pt5Percent = '\002\000\0039',
	WordE112Texture85Percent = '\002\000\003R',
	WordE112Texture87Pt5Percent = '\002\000\003k',
	WordE112Texture90Percent = '\002\000\003\204',
	WordE112Texture92Pt5Percent = '\002\000\003\235',
	WordE112Texture95Percent = '\002\000\003\266',
	WordE112Texture97Pt5Percent = '\002\000\003\317',
	WordE112TextureSolid = '\002\000\003\350',
	WordE112TextureDarkHorizontal = '\001\377\377\377',
	WordE112TextureDarkVertical = '\001\377\377\376',
	WordE112TextureDarkDiagonalDown = '\001\377\377\375',
	WordE112TextureDarkDiagonalUp = '\001\377\377\374',
	WordE112TextureDarkCross = '\001\377\377\373',
	WordE112TextureDarkDiagonalCross = '\001\377\377\372',
	WordE112TextureHorizontal = '\001\377\377\371',
	WordE112TextureVertical = '\001\377\377\370',
	WordE112TextureDiagonalDown = '\001\377\377\367',
	WordE112TextureDiagonalUp = '\001\377\377\366',
	WordE112TextureCross = '\001\377\377\365',
	WordE112TextureDiagonalCross = '\001\377\377\364'
};
typedef enum WordE112 WordE112;

enum WordE113 {
	WordE113UnderlineNone = '\002\001\000\000',
	WordE113UnderlineSingle = '\002\001\000\001',
	WordE113UnderlineWords = '\002\001\000\002',
	WordE113UnderlineDouble = '\002\001\000\003',
	WordE113UnderlineDotted = '\002\001\000\004',
	WordE113UnderlineThick = '\002\001\000\006',
	WordE113UnderlineDash = '\002\001\000\007',
	WordE113UnderlineDotDash = '\002\001\000\011',
	WordE113UnderlineDotDotDash = '\002\001\000\012',
	WordE113UnderlineWavy = '\002\001\000\013',
	WordE113UnderlineWavyHeavy = '\002\001\000\033',
	WordE113UnderlineDottedHeavy = '\002\001\000\024',
	WordE113UnderlineDashHeavy = '\002\001\000\027',
	WordE113UnderlineDotDashHeavy = '\002\001\000\031',
	WordE113UnderlineDotDotDashHeavy = '\002\001\000\032',
	WordE113UnderlineDashLong = '\002\001\000\'',
	WordE113UnderlineDashLongHeavy = '\002\001\0007',
	WordE113UnderlineWavyDouble = '\002\001\000+'
};
typedef enum WordE113 WordE113;

enum WordE114 {
	WordE114EmphasisMarkNone = '\002\002\000\000',
	WordE114EmphasisMarkOverSolidCircle = '\002\002\000\001',
	WordE114EmphasisMarkOverComma = '\002\002\000\002',
	WordE114EmphasisMarkOverWhiteCircle = '\002\002\000\003',
	WordE114EmphasisMarkUnderSolidCircle = '\002\002\000\004'
};
typedef enum WordE114 WordE114;

enum WordE115 {
	WordE115ListSeparator = '\002\003\000\021',
	WordE115DecimalSeparator = '\002\003\000\022',
	WordE115ThousandsSeparator = '\002\003\000\023',
	WordE115CurrencyCode = '\002\003\000\024',
	WordE115TwentyFourHourClock = '\002\003\000\025',
	WordE115InternationalAm = '\002\003\000\026',
	WordE115InternationalPm = '\002\003\000\027',
	WordE115TimeSeparator = '\002\003\000\030',
	WordE115DateSeparator = '\002\003\000\031',
	WordE115ProductLanguageID = '\002\003\000\032'
};
typedef enum WordE115 WordE115;

enum WordE116 {
	WordE116AutoExec = '\002\004\000\000',
	WordE116AutoNew = '\002\004\000\001',
	WordE116AutoOpen = '\002\004\000\002',
	WordE116AutoClose = '\002\004\000\003',
	WordE116AutoExit = '\002\004\000\004'
};
typedef enum WordE116 WordE116;

enum WordE117 {
	WordE117CaptionPositionAbove = '\002\005\000\000',
	WordE117CaptionPositionBelow = '\002\005\000\001'
};
typedef enum WordE117 WordE117;

enum WordE118 {
	WordE118Us = '\002\006\000\001',
	WordE118Canada = '\002\006\000\002',
	WordE118LatinAmerica = '\002\006\000\003',
	WordE118Netherlands = '\002\006\000\037',
	WordE118France = '\002\006\000!',
	WordE118Spain = '\002\006\000\"',
	WordE118Italy = '\002\006\000\'',
	WordE118Uk = '\002\006\000,',
	WordE118Denmark = '\002\006\000-',
	WordE118Sweden = '\002\006\000.',
	WordE118Norway = '\002\006\000/',
	WordE118Germany = '\002\006\0001',
	WordE118Peru = '\002\006\0003',
	WordE118Mexico = '\002\006\0004',
	WordE118Argentina = '\002\006\0006',
	WordE118Brazil = '\002\006\0007',
	WordE118Chile = '\002\006\0008',
	WordE118Venezuela = '\002\006\000:',
	WordE118Japan = '\002\006\000Q',
	WordE118Taiwan = '\002\006\003v',
	WordE118China = '\002\006\000V',
	WordE118Korea = '\002\006\000R',
	WordE118Finland = '\002\006\001f',
	WordE118Iceland = '\002\006\001b'
};
typedef enum WordE118 WordE118;

enum WordE119 {
	WordE119HeadingSeparatorNone = '\002\007\000\000',
	WordE119HeadingSeparatorBlankLine = '\002\007\000\001',
	WordE119HeadingSeparatorLetter = '\002\007\000\002',
	WordE119HeadingSeparatorLetterLow = '\002\007\000\003',
	WordE119HeadingSeparatorLetterFull = '\002\007\000\004'
};
typedef enum WordE119 WordE119;

enum WordE120 {
	WordE120SeparatorHyphen = '\002\010\000\000',
	WordE120SeparatorPeriod = '\002\010\000\001',
	WordE120SeparatorColon = '\002\010\000\002',
	WordE120SeparatorEmDash = '\002\010\000\003',
	WordE120SeparatorEnDash = '\002\010\000\004'
};
typedef enum WordE120 WordE120;

enum WordE121 {
	WordE121AlignPageNumberLeft = '\002\011\000\000',
	WordE121AlignPageNumberCenter = '\002\011\000\001',
	WordE121AlignPageNumberRight = '\002\011\000\002',
	WordE121AlignPageNumberInside = '\002\011\000\003',
	WordE121AlignPageNumberOutside = '\002\011\000\004'
};
typedef enum WordE121 WordE121;

enum WordE122 {
	WordE122BorderTop = '\002\011\377\377',
	WordE122BorderLeft = '\002\011\377\376',
	WordE122BorderBottom = '\002\011\377\375',
	WordE122BorderRight = '\002\011\377\374',
	WordE122BorderHorizontal = '\002\011\377\373',
	WordE122BorderVertical = '\002\011\377\372',
	WordE122BorderDiagonalDown = '\002\011\377\371',
	WordE122BorderDiagonalUp = '\002\011\377\370'
};
typedef enum WordE122 WordE122;

enum WordE123 {
	WordE123FrameTop = '\001\373\275\301',
	WordE123FrameLeft = '\001\373\275\302',
	WordE123FrameBottom = '\001\373\275\303',
	WordE123FrameRight = '\001\373\275\304',
	WordE123FrameCenter = '\001\373\275\305',
	WordE123FrameInside = '\001\373\275\306',
	WordE123FrameOutside = '\001\373\275\307'
};
typedef enum WordE123 WordE123;

enum WordE124 {
	WordE124AnimationNone = '\002\014\000\000',
	WordE124AnimationLasVegasLights = '\002\014\000\001',
	WordE124AnimationBlinkingBackground = '\002\014\000\002',
	WordE124AnimationSparkleText = '\002\014\000\003',
	WordE124AnimationMarchingBlackAnts = '\002\014\000\004',
	WordE124AnimationMarchingRedAnts = '\002\014\000\005',
	WordE124AnimationShimmer = '\002\014\000\006'
};
typedef enum WordE124 WordE124;

enum WordE125 {
	WordE125NextCase = '\002\014\377\377',
	WordE125LowerCase = '\002\015\000\000',
	WordE125UpperCase = '\002\015\000\001',
	WordE125TitleWord = '\002\015\000\002',
	WordE125TitleSentence = '\002\015\000\004',
	WordE125ToggleCase = '\002\015\000\005',
	WordE125CaseHalfWidth = '\002\015\000\006',
	WordE125CaseFullWidth = '\002\015\000\007',
	WordE125CaseKatakana = '\002\015\000\010',
	WordE125CaseHiragana = '\002\015\000\011'
};
typedef enum WordE125 WordE125;

enum WordE127 {
	WordE12710Sentences = '\002\016\377\376',
	WordE12720Sentences = '\002\016\377\375',
	WordE127100Words = '\002\016\377\374',
	WordE127500Words = '\002\016\377\373',
	WordE12710Percent = '\002\016\377\372',
	WordE12725Percent = '\002\016\377\371',
	WordE12750Percent = '\002\016\377\370',
	WordE12775Percent = '\002\016\377\367'
};
typedef enum WordE127 WordE127;

enum WordE128 {
	WordE128StyleTypeParagraph = '\002\020\000\001',
	WordE128StyleTypeCharacter = '\002\020\000\002',
	WordE128StyleTypeTable = '\002\020\000\003',
	WordE128StyleTypeList = '\002\020\000\004'
};
typedef enum WordE128 WordE128;

enum WordE129 {
	WordE129ACharacterItem = '\002\021\000\001',
	WordE129AWordItem = '\002\021\000\002',
	WordE129ASentenceItem = '\002\021\000\003',
	WordE129AParagraphItem = '\002\021\000\004',
	WordE129ALineItem = '\002\021\000\005',
	WordE129AStoryItem = '\002\021\000\006',
	WordE129AScreen = '\002\021\000\007',
	WordE129ASection = '\002\021\000\010',
	WordE129AColumn = '\002\021\000\011',
	WordE129ARow = '\002\021\000\012',
	WordE129AWindow = '\002\021\000\013',
	WordE129ACell = '\002\021\000\014',
	WordE129ACharacterFormatting = '\002\021\000\015',
	WordE129AParagraphFormatting = '\002\021\000\016',
	WordE129ATable = '\002\021\000\017',
	WordE129AnItemUnit = '\002\021\000\020'
};
typedef enum WordE129 WordE129;

enum WordE130 {
	WordE130GotoABookmarkItem = '\002\021\377\377',
	WordE130GotoASectionItem = '\002\022\000\000',
	WordE130GotoAPageItem = '\002\022\000\001',
	WordE130GotoATableItem = '\002\022\000\002',
	WordE130GotoALineItem = '\002\022\000\003',
	WordE130GotoAFootnoteItem = '\002\022\000\004',
	WordE130GotoAnEndnoteItem = '\002\022\000\005',
	WordE130GotoACommentItem = '\002\022\000\006',
	WordE130GotoAFieldItem = '\002\022\000\007',
	WordE130GotoAGraphic = '\002\022\000\010',
	WordE130GotoAnObjectItem = '\002\022\000\011',
	WordE130GotoAnEquation = '\002\022\000\012',
	WordE130GotoAHeadingItem = '\002\022\000\013',
	WordE130GotoAPercent = '\002\022\000\014',
	WordE130GotoASpellingError = '\002\022\000\015',
	WordE130GotoAGrammaticalError = '\002\022\000\016',
	WordE130GotoAProofreadingError = '\002\022\000\017'
};
typedef enum WordE130 WordE130;

enum WordE131 {
	WordE131TheFirstItem = '\002\023\000\001',
	WordE131TheLastItem = '\002\022\377\377',
	WordE131TheNextItem = '\002\023\000\002',
	WordE131Relative = '\002\023\000\002',
	WordE131ThePreviousItem = '\002\023\000\003',
	WordE131Absolute = '\002\023\000\001'
};
typedef enum WordE131 WordE131;

enum WordE132 {
	WordE132CollapseStart = '\002\024\000\001',
	WordE132CollapseEnd = '\002\024\000\000'
};
typedef enum WordE132 WordE132;

enum WordE133 {
	WordE133RowHeightAuto = '\002\025\000\000',
	WordE133RowHeightAtLeast = '\002\025\000\001',
	WordE133RowHeightExactly = '\002\025\000\002'
};
typedef enum WordE133 WordE133;

enum WordE134 {
	WordE134FrameAuto = '\002\026\000\000',
	WordE134FrameAtLeast = '\002\026\000\001',
	WordE134FrameExact = '\002\026\000\002'
};
typedef enum WordE134 WordE134;

enum WordE135 {
	WordE135InsertCellsShiftRight = '\002\027\000\000',
	WordE135InsertCellsShiftDown = '\002\027\000\001',
	WordE135InsertCellsEntireRow = '\002\027\000\002',
	WordE135InsertCellsEntireColumn = '\002\027\000\003'
};
typedef enum WordE135 WordE135;

enum WordE136 {
	WordE136DeleteCellsShiftLeft = '\002\030\000\000',
	WordE136DeleteCellsShiftUp = '\002\030\000\001',
	WordE136DeleteCellsEntireRow = '\002\030\000\002',
	WordE136DeleteCellsEntireColumn = '\002\030\000\003'
};
typedef enum WordE136 WordE136;

enum WordE137 {
	WordE137ListApplyToWholeList = '\002\031\000\000',
	WordE137ListApplyToThisPointForward = '\002\031\000\001',
	WordE137ListApplyToSelection = '\002\031\000\002'
};
typedef enum WordE137 WordE137;

enum WordE138 {
	WordE138AlertsNone = '\002\032\000\000',
	WordE138AlertsMessageBox = '\002\031\377\376',
	WordE138AlertsAll = '\002\031\377\377'
};
typedef enum WordE138 WordE138;

enum WordE139 {
	WordE139CursorWait = '\002\033\000\000',
	WordE139CursorIbeam = '\002\033\000\001',
	WordE139CursorNormal = '\002\033\000\002',
	WordE139CursorNorthwestArrow = '\002\033\000\003'
};
typedef enum WordE139 WordE139;

enum WordE141 {
	WordE141AdjustNone = '\002\035\000\000',
	WordE141AdjustProportional = '\002\035\000\001',
	WordE141AdjustFirstColumn = '\002\035\000\002',
	WordE141AdjustSameWidth = '\002\035\000\003'
};
typedef enum WordE141 WordE141;

enum WordE142 {
	WordE142AlignParagraphLeft = '\002\036\000\000',
	WordE142AlignParagraphCenter = '\002\036\000\001',
	WordE142AlignParagraphRight = '\002\036\000\002',
	WordE142AlignParagraphJustify = '\002\036\000\003',
	WordE142AlignParagraphDistribute = '\002\036\000\004'
};
typedef enum WordE142 WordE142;

enum WordE143 {
	WordE143ListLevelAlignLeft = '\002\037\000\000',
	WordE143ListLevelAlignCenter = '\002\037\000\001',
	WordE143ListLevelAlignRight = '\002\037\000\002'
};
typedef enum WordE143 WordE143;

enum WordE144 {
	WordE144AlignRowLeft = '\002 \000\000',
	WordE144AlignRowCenter = '\002 \000\001',
	WordE144AlignRowRight = '\002 \000\002'
};
typedef enum WordE144 WordE144;

enum WordE145 {
	WordE145AlignTabLeft = '\002!\000\000',
	WordE145AlignTabCenter = '\002!\000\001',
	WordE145AlignTabRight = '\002!\000\002',
	WordE145AlignTabDecimal = '\002!\000\003',
	WordE145AlignTabBar = '\002!\000\004',
	WordE145AlignTabList = '\002!\000\006'
};
typedef enum WordE145 WordE145;

enum WordE146 {
	WordE146AlignVerticalTop = '\002\"\000\000',
	WordE146AlignVerticalCenter = '\002\"\000\001',
	WordE146AlignVerticalJustify = '\002\"\000\002',
	WordE146AlignVerticalBottom = '\002\"\000\003'
};
typedef enum WordE146 WordE146;

enum WordE147 {
	WordE147CellAlignVerticalTop = '\002#\000\000',
	WordE147CellAlignVerticalCenter = '\002#\000\001',
	WordE147CellAlignVerticalBottom = '\002#\000\003'
};
typedef enum WordE147 WordE147;

enum WordE148 {
	WordE148RangeAnnotationAlignmentCenter = '\002$\000\000',
	WordE148ZeroOneZero = '\002$\000\001',
	WordE148OneTwoOne = '\002$\000\002',
	WordE148RangeAnnotationAlignmentLeft = '\002$\000\003',
	WordE148RangeAnnotationAlignmentRight = '\002$\000\004'
};
typedef enum WordE148 WordE148;

enum WordE149 {
	WordE149TrailingTab = '\002%\000\000',
	WordE149TrailingSpace = '\002%\000\001',
	WordE149TrailingNone = '\002%\000\002'
};
typedef enum WordE149 WordE149;

enum WordE150 {
	WordE150BulletGallery = '\002&\000\001',
	WordE150NumberGallery = '\002&\000\002',
	WordE150OutlineNumberGallery = '\002&\000\003'
};
typedef enum WordE150 WordE150;

enum WordE151 {
	WordE151ListNumberStyleArabic = '\002\'\000\000',
	WordE151ListNumberStyleUppercaseRoman = '\002\'\000\001',
	WordE151ListNumberStyleLowercaseRoman = '\002\'\000\002',
	WordE151ListNumberStyleUppercaseLetter = '\002\'\000\003',
	WordE151ListNumberStyleLowercaseLetter = '\002\'\000\004',
	WordE151ListNumberStyleOrdinal = '\002\'\000\005',
	WordE151ListNumberStyleCardinalText = '\002\'\000\006',
	WordE151ListNumberStyleOrdinalText = '\002\'\000\007',
	WordE151ListNumberStyleKanji = '\002\'\000\012',
	WordE151ListNumberStyleKanjiDigit = '\002\'\000\013',
	WordE151ListNumberStyleAiueoHalfWidth = '\002\'\000\014',
	WordE151ListNumberStyleIrohaHalfWidth = '\002\'\000\015',
	WordE151ListNumberStyleArabicFullWidth = '\002\'\000\016',
	WordE151ListNumberStyleKanjiTraditional = '\002\'\000\020',
	WordE151ListNumberStyleKanjiTraditional2 = '\002\'\000\021',
	WordE151ListNumberStyleNumberInCircle = '\002\'\000\022',
	WordE151ListNumberStyleAiueo = '\002\'\000\024',
	WordE151ListNumberStyleIroha = '\002\'\000\025',
	WordE151ListNumberStyleArabicLz = '\002\'\000\026',
	WordE151ListNumberStyleBullet = '\002\'\000\027',
	WordE151ListNumberStyleGanada = '\002\'\000\030',
	WordE151ListNumberStyleChosung = '\002\'\000\031',
	WordE151ListNumberStyleGbnum1 = '\002\'\000\032',
	WordE151ListNumberStyleGbnum2 = '\002\'\000\033',
	WordE151ListNumberStyleGbnum3 = '\002\'\000\034',
	WordE151ListNumberStyleGbnum4 = '\002\'\000\035',
	WordE151ListNumberStyleZodiac1 = '\002\'\000\036',
	WordE151ListNumberStyleZodiac2 = '\002\'\000\037',
	WordE151ListNumberStyleZodiac3 = '\002\'\000 ',
	WordE151ListNumberStyleTradChinNum1 = '\002\'\000!',
	WordE151ListNumberStyleTradChinNum2 = '\002\'\000\"',
	WordE151ListNumberStyleTradChinNum3 = '\002\'\000#',
	WordE151ListNumberStyleTradChinNum4 = '\002\'\000$',
	WordE151ListNumberStyleSimpChinNum1 = '\002\'\000%',
	WordE151ListNumberStyleSimpChinNum2 = '\002\'\000&',
	WordE151ListNumberStyleSimpChinNum3 = '\002\'\000\'',
	WordE151ListNumberStyleSimpChinNum4 = '\002\'\000(',
	WordE151ListNumberStyleHanjaRead = '\002\'\000)',
	WordE151ListNumberStyleHanjaReadDigit = '\002\'\000*',
	WordE151ListNumberStyleHangul = '\002\'\000+',
	WordE151ListNumberStyleHanja = '\002\'\000,',
	WordE151ListNumberStylePictureBullet = '\002\'\000\371',
	WordE151ListNumberStyleLegal = '\002\'\000\375',
	WordE151ListNumberStyleLegalLz = '\002\'\000\376',
	WordE151ListNumberStyleNone = '\002\'\000\377'
};
typedef enum WordE151 WordE151;

enum WordE152 {
	WordE152NoteNumberStyleArabic = '\002(\000\000',
	WordE152NoteNumberStyleUppercaseRoman = '\002(\000\001',
	WordE152NoteNumberStyleLowercaseRoman = '\002(\000\002',
	WordE152NoteNumberStyleUppercaseLetter = '\002(\000\003',
	WordE152NoteNumberStyleLowercaseLetter = '\002(\000\004',
	WordE152NoteNumberStyleSymbol = '\002(\000\011',
	WordE152NoteNumberStyleArabicFullWidth = '\002(\000\016',
	WordE152NoteNumberStyleKanji = '\002(\000\012',
	WordE152NoteNumberStyleKanjiDigit = '\002(\000\013',
	WordE152NoteNumberStyleKanjiTraditional = '\002(\000\020',
	WordE152NoteNumberStyleNumberInCircle = '\002(\000\022',
	WordE152NoteNumberStyleHanjaRead = '\002(\000)',
	WordE152NoteNumberStyleHanjaReadDigit = '\002(\000*',
	WordE152NoteNumberStyleTradChinNum1 = '\002(\000!',
	WordE152NoteNumberStyleTradChinNum2 = '\002(\000\"',
	WordE152NoteNumberStyleSimpChinNum1 = '\002(\000%',
	WordE152NoteNumberStyleSimpChinNum2 = '\002(\000&'
};
typedef enum WordE152 WordE152;

enum WordE153 {
	WordE153CaptionNumberStyleArabic = '\002)\000\000',
	WordE153CaptionNumberStyleUppercaseRoman = '\002)\000\001',
	WordE153CaptionNumberStyleLowercaseRoman = '\002)\000\002',
	WordE153CaptionNumberStyleUppercaseLetter = '\002)\000\003',
	WordE153CaptionNumberStyleLowercaseLetter = '\002)\000\004',
	WordE153CaptionNumberStyleArabicFullWidth = '\002)\000\016',
	WordE153CaptionNumberStyleKanji = '\002)\000\012',
	WordE153CaptionNumberStyleKanjiDigit = '\002)\000\013',
	WordE153CaptionNumberStyleKanjiTraditional = '\002)\000\020',
	WordE153CaptionNumberStyleNumberInCircle = '\002)\000\022',
	WordE153CaptionNumberStyleGanada = '\002)\000\030',
	WordE153CaptionNumberStyleChosung = '\002)\000\031',
	WordE153CaptionNumberStyleZodiac1 = '\002)\000\036',
	WordE153CaptionNumberStyleZodiac2 = '\002)\000\037',
	WordE153CaptionNumberStyleHanjaRead = '\002)\000)',
	WordE153CaptionNumberStyleHanjaReadDigit = '\002)\000*',
	WordE153CaptionNumberStyleTradChinNum2 = '\002)\000\"',
	WordE153CaptionNumberStyleTradChinNum3 = '\002)\000#',
	WordE153CaptionNumberStyleSimpChinNum2 = '\002)\000&',
	WordE153CaptionNumberStyleSimpChinNum3 = '\002)\000\''
};
typedef enum WordE153 WordE153;

enum WordE154 {
	WordE154PageNumberStyleArabic = '\002*\000\000',
	WordE154PageNumberStyleUppercaseRoman = '\002*\000\001',
	WordE154PageNumberStyleLowercaseRoman = '\002*\000\002',
	WordE154PageNumberStyleUppercaseLetter = '\002*\000\003',
	WordE154PageNumberStyleLowercaseLetter = '\002*\000\004',
	WordE154PageNumberStyleArabicFullWidth = '\002*\000\016',
	WordE154PageNumberStyleKanji = '\002*\000\012',
	WordE154PageNumberStyleKanjiDigit = '\002*\000\013',
	WordE154PageNumberStyleKanjiTraditional = '\002*\000\020',
	WordE154PageNumberStyleNumberInCircle = '\002*\000\022',
	WordE154PageNumberStyleHanjaRead = '\002*\000)',
	WordE154PageNumberStyleHanjaReadDigit = '\002*\000*',
	WordE154PageNumberStyleTradChinNum1 = '\002*\000!',
	WordE154PageNumberStyleTradChinNum2 = '\002*\000\"',
	WordE154PageNumberStyleSimpChinNum1 = '\002*\000%',
	WordE154PageNumberStyleSimpChinNum2 = '\002*\000&'
};
typedef enum WordE154 WordE154;

enum WordE155 {
	WordE155StatisticWords = '\002+\000\000',
	WordE155StatisticLines = '\002+\000\001',
	WordE155StatisticPages = '\002+\000\002',
	WordE155StatisticCharacters = '\002+\000\003',
	WordE155StatisticParagraphs = '\002+\000\004',
	WordE155StatisticCharactersWithSpaces = '\002+\000\005',
	WordE155StatisticEastAsianCharacters = '\002+\000\006'
};
typedef enum WordE155 WordE155;

enum WordE156 {
	WordE156PropertyTitle = '\002,\000\001',
	WordE156PropertySubject = '\002,\000\002',
	WordE156PropertyAuthor = '\002,\000\003',
	WordE156PropertyKeywords = '\002,\000\004',
	WordE156PropertyComments = '\002,\000\005',
	WordE156PropertyTemplate = '\002,\000\006',
	WordE156PropertyLastAuthor = '\002,\000\007',
	WordE156PropertyRevision = '\002,\000\010',
	WordE156PropertyAppName = '\002,\000\011',
	WordE156PropertyTimeLastPrinted = '\002,\000\012',
	WordE156PropertyTimeCreated = '\002,\000\013',
	WordE156PropertyTimeLastSaved = '\002,\000\014',
	WordE156PropertyVbatotalEdit = '\002,\000\015',
	WordE156PropertyPages = '\002,\000\016',
	WordE156PropertyWords = '\002,\000\017',
	WordE156PropertyCharacters = '\002,\000\020',
	WordE156PropertySecurity = '\002,\000\021',
	WordE156PropertyCategory = '\002,\000\022',
	WordE156PropertyFormat = '\002,\000\023',
	WordE156PropertyManager = '\002,\000\024',
	WordE156PropertyCompany = '\002,\000\025',
	WordE156PropertyBytes = '\002,\000\026',
	WordE156PropertyLines = '\002,\000\027',
	WordE156PropertyParas = '\002,\000\030',
	WordE156PropertySlides = '\002,\000\031',
	WordE156PropertyNotes = '\002,\000\032',
	WordE156PropertyHiddenSlides = '\002,\000\033',
	WordE156PropertyMmclips = '\002,\000\034',
	WordE156PropertyHyperlinkBase = '\002,\000\035',
	WordE156PropertyCharsWspaces = '\002,\000\036'
};
typedef enum WordE156 WordE156;

enum WordE157 {
	WordE157LineSpaceSingle = '\002-\000\000',
	WordE157LineSpace1Pt5 = '\002-\000\001',
	WordE157LineSpaceDouble = '\002-\000\002',
	WordE157LineSpaceAtLeast = '\002-\000\003',
	WordE157LineSpaceExactly = '\002-\000\004',
	WordE157LineSpaceMultiple = '\002-\000\005'
};
typedef enum WordE157 WordE157;

enum WordE158 {
	WordE158NumberParagraph = '\002.\000\001',
	WordE158NumberListnum = '\002.\000\002',
	WordE158NumberAllNumbers = '\002.\000\003'
};
typedef enum WordE158 WordE158;

enum WordE159 {
	WordE159ListNoNumbering = '\002/\000\000',
	WordE159ListListnumOnly = '\002/\000\001',
	WordE159ListBullet = '\002/\000\002',
	WordE159ListSimpleNumbering = '\002/\000\003',
	WordE159ListOutlineNumbering = '\002/\000\004',
	WordE159ListMixedNumbering = '\002/\000\005',
	WordE159ListPictureBullet = '\002/\000\006'
};
typedef enum WordE159 WordE159;

enum WordE160 {
	WordE160MainTextStory = '\0020\000\001',
	WordE160FootnotesStory = '\0020\000\002',
	WordE160EndnotesStory = '\0020\000\003',
	WordE160CommentsStory = '\0020\000\004',
	WordE160TextFrameStory = '\0020\000\005',
	WordE160EvenPagesHeaderStory = '\0020\000\006',
	WordE160PrimaryHeaderStory = '\0020\000\007',
	WordE160EvenPagesFooterStory = '\0020\000\010',
	WordE160PrimaryFooterStory = '\0020\000\011',
	WordE160FirstPageHeaderStory = '\0020\000\012',
	WordE160FirstPageFooterStory = '\0020\000\013',
	WordE160FootnoteSeparatorStory = '\0020\000\014',
	WordE160FootnoteContinuationSeparatorStory = '\0020\000\015',
	WordE160FootnoteContinuationNoticeStory = '\0020\000\016',
	WordE160EndnoteSeparatorStory = '\0020\000\017',
	WordE160EndnoteContinuationSeparatorStory = '\0020\000\020',
	WordE160EndnoteContinuationNoticeStory = '\0020\000\021'
};
typedef enum WordE160 WordE160;

enum WordE161 {
	WordE161FormatDocument97 = '\0021\000\000',
	WordE161FormatTemplate97 = '\0021\000\001',
	WordE161FormatText = '\0021\000\002',
	WordE161FormatTextLineBreaks = '\0021\000\003',
	WordE161FormatDostext = '\0021\000\004',
	WordE161FormatDostextLineBreaks = '\0021\000\005',
	WordE161FormatRtf = '\0021\000\006',
	WordE161FormatUnicodeText = '\0021\000\007',
	WordE161FormatHTML = '\0021\000\010',
	WordE161FormatWebArchive = '\0021\000\011',
	WordE161FormatStationery = '\0021\000\012',
	WordE161FormatXml = '\0021\000\013',
	WordE161FormatDocument = '\0021\000\014',
	WordE161FormatDocumentME = '\0021\000\015',
	WordE161FormatTemplate = '\0021\000\016',
	WordE161FormatTemplateME = '\0021\000\017',
	WordE161FormatPDF = '\0021\000\020',
	WordE161FormatFlatDocument = '\0021\000\021',
	WordE161FormatFlatDocumentME = '\0021\000\022',
	WordE161FormatFlatTemplate = '\0021\000\023',
	WordE161FormatFlatTemplateME = '\0021\000\024',
	WordE161FormatCustomDictionary = '\0021\000\025',
	WordE161FormatExcludeDictionary = '\0021\000\026',
	WordE161FormatDocumentAuto = '\0021\015\014',
	WordE161FormatTemplateAuto = '\0021\015\007'
};
typedef enum WordE161 WordE161;

enum WordE162 {
	WordE162OpenFormatAuto = '\0022\000\000',
	WordE162OpenFormatDocument = '\0022\000\001',
	WordE162OpenFormatTemplate = '\0022\000\002',
	WordE162OpenFormatRtf = '\0022\000\003',
	WordE162OpenFormatText = '\0022\000\004',
	WordE162OpenFormatUnicodeText = '\0022\000\005',
	WordE162OpenFormatEncodedText = '\0022\000\005',
	WordE162OpenFormatMacReadable = '\0022\000\006',
	WordE162OpenFormatWebPages = '\0022\000\007',
	WordE162OpenFormatXml = '\0022\000\010',
	WordE162OpenFormatDocument97 = '\0022\000\011',
	WordE162OpenFormatTemplate97 = '\0022\000\012',
	WordE162OpenFormatOffice = '\0022\000\013'
};
typedef enum WordE162 WordE162;

enum WordE163 {
	WordE163HeaderFooterPrimary = '\0023\000\001',
	WordE163HeaderFooterFirstPage = '\0023\000\002',
	WordE163HeaderFooterEvenPages = '\0023\000\003'
};
typedef enum WordE163 WordE163;

enum WordE164 {
	WordE164Toctemplate = '\0024\000\000',
	WordE164Tocclassic = '\0024\000\001',
	WordE164Tocdistinctive = '\0024\000\002',
	WordE164Tocfancy = '\0024\000\003',
	WordE164Tocmodern = '\0024\000\004',
	WordE164Tocformal = '\0024\000\005',
	WordE164Tocsimple = '\0024\000\006'
};
typedef enum WordE164 WordE164;

enum WordE165 {
	WordE165Toftemplate = '\0025\000\000',
	WordE165Tofclassic = '\0025\000\001',
	WordE165Tofdistinctive = '\0025\000\002',
	WordE165Tofcentered = '\0025\000\003',
	WordE165Tofformal = '\0025\000\004',
	WordE165Tofsimple = '\0025\000\005'
};
typedef enum WordE165 WordE165;

enum WordE166 {
	WordE166Toatemplate = '\0026\000\000',
	WordE166Toaclassic = '\0026\000\001',
	WordE166Toadistinctive = '\0026\000\002',
	WordE166Toaformal = '\0026\000\003',
	WordE166Toasimple = '\0026\000\004'
};
typedef enum WordE166 WordE166;

enum WordE167 {
	WordE167LineStyleNone = '\0027\000\000',
	WordE167LineStyleSingle = '\0027\000\001',
	WordE167LineStyleDot = '\0027\000\002',
	WordE167LineStyleDashSmallGap = '\0027\000\003',
	WordE167LineStyleDashLargeGap = '\0027\000\004',
	WordE167LineStyleDashDot = '\0027\000\005',
	WordE167LineStyleDashDotDot = '\0027\000\006',
	WordE167LineStyleDouble = '\0027\000\007',
	WordE167LineStyleTriple = '\0027\000\010',
	WordE167LineStyleThinThickSmallGap = '\0027\000\011',
	WordE167LineStyleThickThinSmallGap = '\0027\000\012',
	WordE167LineStyleThinThickThinSmallGap = '\0027\000\013',
	WordE167LineStyleThinThickMedGap = '\0027\000\014',
	WordE167LineStyleThickThinMedGap = '\0027\000\015',
	WordE167LineStyleThinThickThinMedGap = '\0027\000\016',
	WordE167LineStyleThinThickLargeGap = '\0027\000\017',
	WordE167LineStyleThickThinLargeGap = '\0027\000\020',
	WordE167LineStyleThinThickThinLargeGap = '\0027\000\021',
	WordE167LineStyleSingleWavy = '\0027\000\022',
	WordE167LineStyleDoubleWavy = '\0027\000\023',
	WordE167LineStyleDashDotStroked = '\0027\000\024',
	WordE167LineStyleEmboss_3D = '\0027\000\025',
	WordE167LineStyleEngrave_3D = '\0027\000\026',
	WordE167LineStyleOutset = '\0027\000\027',
	WordE167LineStyleInset = '\0027\000\030'
};
typedef enum WordE167 WordE167;

enum WordE168 {
	WordE168LineWidth25Point = '\0028\000\002',
	WordE168LineWidth50Point = '\0028\000\004',
	WordE168LineWidth75Point = '\0028\000\006',
	WordE168LineWidth100Point = '\0028\000\010',
	WordE168LineWidth150Point = '\0028\000\014',
	WordE168LineWidth225Point = '\0028\000\022',
	WordE168LineWidth300Point = '\0028\000\030',
	WordE168LineWidth450Point = '\0028\000$',
	WordE168LineWidth600Point = '\0028\0000'
};
typedef enum WordE168 WordE168;

enum WordE169 {
	WordE169SectionBreakNextPage = '\0029\000\002',
	WordE169SectionBreakContinuous = '\0029\000\003',
	WordE169SectionBreakEvenPage = '\0029\000\004',
	WordE169SectionBreakOddPage = '\0029\000\005',
	WordE169LineBreak = '\0029\000\006',
	WordE169PageBreak = '\0029\000\007',
	WordE169ColumnBreak = '\0029\000\010'
};
typedef enum WordE169 WordE169;

enum WordE313 {
	WordE313ContinuedMaster = '\002\311\000\000'
};
typedef enum WordE313 WordE313;

enum WordE170 {
	WordE170TabLeaderSpaces = '\002:\000\000',
	WordE170TabLeaderDots = '\002:\000\001',
	WordE170TabLeaderDashes = '\002:\000\002',
	WordE170TabLeaderLines = '\002:\000\003',
	WordE170TabLeaderHeavy = '\002:\000\004',
	WordE170TabLeaderMiddleDot = '\002:\000\005'
};
typedef enum WordE170 WordE170;

enum WordE171 {
	WordE171Inches = '\002;\000\000',
	WordE171Centimeters = '\002;\000\001',
	WordE171Millimeters = '\002;\000\002',
	WordE171Points = '\002;\000\003',
	WordE171Picas = '\002;\000\004'
};
typedef enum WordE171 WordE171;

enum WordE172 {
	WordE172DropNone = '\002<\000\000',
	WordE172DropNormal = '\002<\000\001',
	WordE172DropMargin = '\002<\000\002'
};
typedef enum WordE172 WordE172;

enum WordE173 {
	WordE173RestartContinuous = '\002=\000\000',
	WordE173RestartSection = '\002=\000\001',
	WordE173RestartPage = '\002=\000\002'
};
typedef enum WordE173 WordE173;

enum WordE174 {
	WordE174BottomOfPage = '\002>\000\000',
	WordE174BeneathText = '\002>\000\001'
};
typedef enum WordE174 WordE174;

enum WordE175 {
	WordE175End_of_section = '\002\?\000\000',
	WordE175End_of_document = '\002\?\000\001'
};
typedef enum WordE175 WordE175;

enum WordE176 {
	WordE176SortSeparateByTabs = '\002@\000\000',
	WordE176SortSeparateByCommas = '\002@\000\001',
	WordE176SortSeparateByDefaultTableSeparator = '\002@\000\002'
};
typedef enum WordE176 WordE176;

enum WordE177 {
	WordE177SeparateByParagraphs = '\002A\000\000',
	WordE177SeparateByTabs = '\002A\000\001',
	WordE177SeparateByCommas = '\002A\000\002',
	WordE177SeparateByDefaultListSeparator = '\002A\000\003'
};
typedef enum WordE177 WordE177;

enum WordE178 {
	WordE178SortFieldAlphanumeric = '\002B\000\000',
	WordE178SortFieldNumeric = '\002B\000\001',
	WordE178SortFieldDate = '\002B\000\002',
	WordE178SortFieldSyllable = '\002B\000\003',
	WordE178SortFieldJapanJis = '\002B\000\004',
	WordE178SortFieldStroke = '\002B\000\005',
	WordE178SortFieldKoreaKs = '\002B\000\006'
};
typedef enum WordE178 WordE178;

enum WordE179 {
	WordE179SortOrderAscending = '\002C\000\000',
	WordE179SortOrderDescending = '\002C\000\001'
};
typedef enum WordE179 WordE179;

enum WordE180 {
	WordE180TableFormatNone = '\002D\000\000',
	WordE180TableFormatSimple1 = '\002D\000\001',
	WordE180TableFormatSimple2 = '\002D\000\002',
	WordE180TableFormatSimple3 = '\002D\000\003',
	WordE180TableFormatClassic1 = '\002D\000\004',
	WordE180TableFormatClassic2 = '\002D\000\005',
	WordE180TableFormatClassic3 = '\002D\000\006',
	WordE180TableFormatClassic4 = '\002D\000\007',
	WordE180TableFormatColorful1 = '\002D\000\010',
	WordE180TableFormatColorful2 = '\002D\000\011',
	WordE180TableFormatColorful3 = '\002D\000\012',
	WordE180TableFormatColumns1 = '\002D\000\013',
	WordE180TableFormatColumns2 = '\002D\000\014',
	WordE180TableFormatColumns3 = '\002D\000\015',
	WordE180TableFormatColumns4 = '\002D\000\016',
	WordE180TableFormatColumns5 = '\002D\000\017',
	WordE180TableFormatGrid1 = '\002D\000\020',
	WordE180TableFormatGrid2 = '\002D\000\021',
	WordE180TableFormatGrid3 = '\002D\000\022',
	WordE180TableFormatGrid4 = '\002D\000\023',
	WordE180TableFormatGrid5 = '\002D\000\024',
	WordE180TableFormatGrid6 = '\002D\000\025',
	WordE180TableFormatGrid7 = '\002D\000\026',
	WordE180TableFormatGrid8 = '\002D\000\027',
	WordE180TableFormatList1 = '\002D\000\030',
	WordE180TableFormatList2 = '\002D\000\031',
	WordE180TableFormatList3 = '\002D\000\032',
	WordE180TableFormatList4 = '\002D\000\033',
	WordE180TableFormatList5 = '\002D\000\034',
	WordE180TableFormatList6 = '\002D\000\035',
	WordE180TableFormatList7 = '\002D\000\036',
	WordE180TableFormatList8 = '\002D\000\037',
	WordE180TableFormat3DEffects1 = '\002D\000 ',
	WordE180TableFormat3DEffects2 = '\002D\000!',
	WordE180TableFormat3DEffects3 = '\002D\000\"',
	WordE180TableFormatContemporary = '\002D\000#',
	WordE180TableFormatElegant = '\002D\000$',
	WordE180TableFormatProfessional = '\002D\000%',
	WordE180TableFormatSubtle1 = '\002D\000&',
	WordE180TableFormatSubtle2 = '\002D\000\'',
	WordE180TableFormatWeb1 = '\002D\000(',
	WordE180TableFormatWeb2 = '\002D\000)',
	WordE180TableFormatWeb3 = '\002D\000*'
};
typedef enum WordE180 WordE180;

enum WordE181 {
	WordE181TableFormatApplyBorders = '\002E\000\001',
	WordE181TableFormatApplyShading = '\002E\000\002',
	WordE181TableFormatApplyFont = '\002E\000\004',
	WordE181TableFormatApplyColor = '\002E\000\010',
	WordE181TableFormatApplyAutoFit = '\002E\000\020',
	WordE181TableFormatApplyHeadingRows = '\002E\000 ',
	WordE181TableFormatApplyLastRow = '\002E\000@',
	WordE181TableFormatApplyFirstColumn = '\002E\000\200',
	WordE181TableFormatApplyLastColumn = '\002E\001\000'
};
typedef enum WordE181 WordE181;

enum WordE182 {
	WordE182LanguageNone = '\002F\000\000',
	WordE182LanguageNoProofing = '\002F\004\000',
	WordE182Danish = '\002F\004\006',
	WordE182German = '\002F\004\007',
	WordE182SwissGerman = '\002F\010\007',
	WordE182AustrianGerman = '\002F\014\007',
	WordE182EnglishAus = '\002F\014\011',
	WordE182EnglishUk = '\002F\010\011',
	WordE182EnglishUs = '\002F\004\011',
	WordE182EnglishCanadian = '\002F\020\011',
	WordE182EnglishNewZealand = '\002F\024\011',
	WordE182EnglishSouthAfrica = '\002F\034\011',
	WordE182Spanish = '\002F\004\012',
	WordE182French = '\002F\004\014',
	WordE182FrenchCanadian = '\002F\014\014',
	WordE182Italian = '\002F\004\020',
	WordE182Dutch = '\002F\004\023',
	WordE182NorwegianBokmol = '\002F\004\024',
	WordE182NorwegianNynorsk = '\002F\010\024',
	WordE182BrazilianPortuguese = '\002F\004\026',
	WordE182Portuguese = '\002F\010\026',
	WordE182Finnish = '\002F\004\013',
	WordE182Swedish = '\002F\004\035',
	WordE182Catalan = '\002F\004\003',
	WordE182Greek = '\002F\004\010',
	WordE182Turkish = '\002F\004\037',
	WordE182Russian = '\002F\004\031',
	WordE182Czech = '\002F\004\005',
	WordE182Hungarian = '\002F\004\016',
	WordE182Polish = '\002F\004\025',
	WordE182Slovenian = '\002F\004$',
	WordE182Basque = '\002F\004-',
	WordE182Malaysian = '\002F\004>',
	WordE182Japanese = '\002F\004\021',
	WordE182Korean = '\002F\004\022',
	WordE182SimplifiedChinese = '\002F\010\004',
	WordE182TraditionalChinese = '\002F\004\004',
	WordE182SwissFrench = '\002F\020\014',
	WordE182Sesotho = '\002F\0040',
	WordE182Tsonga = '\002F\0041',
	WordE182Tswana = '\002F\0042',
	WordE182Venda = '\002F\0043',
	WordE182Xhosa = '\002F\0044',
	WordE182Zulu = '\002F\0045',
	WordE182Afrikaans = '\002F\0046',
	WordE182Arabic = '\002F\004\001',
	WordE182Hebrew = '\002F\004\015',
	WordE182Slovak = '\002F\004\033',
	WordE182Farsi = '\002F\004)',
	WordE182Romanian = '\002F\004\030',
	WordE182Croatian = '\002F\004\032',
	WordE182Ukrainian = '\002F\004\"',
	WordE182Byelorussian = '\002F\004#',
	WordE182Estonian = '\002F\004%',
	WordE182Latvian = '\002F\004&',
	WordE182Macedonian = '\002F\004/',
	WordE182SerbianLatin = '\002F\010\032',
	WordE182SerbianCyrillic = '\002F\014\032',
	WordE182Icelandic = '\002F\004\017',
	WordE182BelgianFrench = '\002F\010\014',
	WordE182BelgianDutch = '\002F\010\023',
	WordE182Bulgarian = '\002F\004\002',
	WordE182MexicanSpanish = '\002F\010\012',
	WordE182SpanishModernSort = '\002F\014\012',
	WordE182SwissItalian = '\002F\010\020'
};
typedef enum WordE182 WordE182;

enum WordE183 {
	WordE183FieldEmpty = '\002F\377\377',
	WordE183FieldRef = '\002G\000\003',
	WordE183FieldIndexEntry = '\002G\000\004',
	WordE183FieldFootnoteRef = '\002G\000\005',
	WordE183FieldSet = '\002G\000\006',
	WordE183FieldIf = '\002G\000\007',
	WordE183FieldIndex = '\002G\000\010',
	WordE183FieldTocEntry = '\002G\000\011',
	WordE183FieldStyleRef = '\002G\000\012',
	WordE183FieldRefDoc = '\002G\000\013',
	WordE183FieldSequence = '\002G\000\014',
	WordE183FieldToc = '\002G\000\015',
	WordE183FieldInfo = '\002G\000\016',
	WordE183FieldTitle = '\002G\000\017',
	WordE183FieldSubject = '\002G\000\020',
	WordE183FieldAuthor = '\002G\000\021',
	WordE183FieldKeyWord = '\002G\000\022',
	WordE183FieldComments = '\002G\000\023',
	WordE183FieldLastSavedBy = '\002G\000\024',
	WordE183FieldCreateDate = '\002G\000\025',
	WordE183FieldSaveDate = '\002G\000\026',
	WordE183FieldPrintDate = '\002G\000\027',
	WordE183FieldRevisionNum = '\002G\000\030',
	WordE183FieldEditTime = '\002G\000\031',
	WordE183FieldNumPages = '\002G\000\032',
	WordE183FieldNumWords = '\002G\000\033',
	WordE183FieldNumChars = '\002G\000\034',
	WordE183FieldFileName = '\002G\000\035',
	WordE183FieldTemplate = '\002G\000\036',
	WordE183FieldDate = '\002G\000\037',
	WordE183FieldTime = '\002G\000 ',
	WordE183FieldPage = '\002G\000!',
	WordE183FieldExpression = '\002G\000\"',
	WordE183FieldQuote = '\002G\000#',
	WordE183FieldInclude = '\002G\000$',
	WordE183FieldPageRef = '\002G\000%',
	WordE183FieldAsk = '\002G\000&',
	WordE183FieldFillIn = '\002G\000\'',
	WordE183FieldData = '\002G\000(',
	WordE183FieldNext = '\002G\000)',
	WordE183FieldNextIf = '\002G\000*',
	WordE183FieldSkipIf = '\002G\000+',
	WordE183FieldMergeRec = '\002G\000,',
	WordE183FieldDde = '\002G\000-',
	WordE183FieldDdeauto = '\002G\000.',
	WordE183FieldGlossary = '\002G\000/',
	WordE183FieldPrint = '\002G\0000',
	WordE183FieldFormula = '\002G\0001',
	WordE183FieldGoToButton = '\002G\0002',
	WordE183Field02470033 = '\002G\0003',
	WordE183FieldAutoNumOutline = '\002G\0004',
	WordE183FieldAutoNumLegal = '\002G\0005',
	WordE183FieldAutoNum = '\002G\0006',
	WordE183FieldImport = '\002G\0007',
	WordE183FieldLink = '\002G\0008',
	WordE183FieldSymbol = '\002G\0009',
	WordE183FieldEmbed = '\002G\000:',
	WordE183FieldMergeField = '\002G\000;',
	WordE183FieldUserName = '\002G\000<',
	WordE183FieldUserInitials = '\002G\000=',
	WordE183FieldUserAddress = '\002G\000>',
	WordE183FieldBarCode = '\002G\000\?',
	WordE183FieldDocVariable = '\002G\000@',
	WordE183FieldSection = '\002G\000A',
	WordE183FieldSectionPages = '\002G\000B',
	WordE183FieldIncludePicture = '\002G\000C',
	WordE183FieldIncludeText = '\002G\000D',
	WordE183FieldFileSize = '\002G\000E',
	WordE183FieldFormTextInput = '\002G\000F',
	WordE183FieldFormCheckBox = '\002G\000G',
	WordE183FieldNoteRef = '\002G\000H',
	WordE183FieldToa = '\002G\000I',
	WordE183FieldToaentry = '\002G\000J',
	WordE183FieldMergeSeq = '\002G\000K',
	WordE183FieldPrivate = '\002G\000M',
	WordE183FieldDatabase = '\002G\000N',
	WordE183FieldAutoText = '\002G\000O',
	WordE183FieldCompare = '\002G\000P',
	WordE183FieldAddin = '\002G\000Q',
	WordE183FieldSubscriber = '\002G\000R',
	WordE183FieldFormDropDown = '\002G\000S',
	WordE183FieldAdvance = '\002G\000T',
	WordE183FieldDocProperty = '\002G\000U',
	WordE183FieldOcx = '\002G\000W',
	WordE183FieldHyperlink = '\002G\000X',
	WordE183FieldAutoTextList = '\002G\000Y',
	WordE183FieldListnum = '\002G\000Z',
	WordE183FieldHtmlactiveX = '\002G\000[',
	WordE183FieldContact = '\002G\000b',
	WordE183FieldUserProperty = '\002G\000c'
};
typedef enum WordE183 WordE183;

enum WordE184 {
	WordE184StyleNormal = '\002G\377\377',
	WordE184StyleEnvelopeAddress = '\002G\377\333',
	WordE184StyleEnvelopeReturn = '\002G\377\332',
	WordE184StyleBodyText = '\002G\377\275',
	WordE184StyleHeading1 = '\002G\377\376',
	WordE184StyleHeading2 = '\002G\377\375',
	WordE184StyleHeading3 = '\002G\377\374',
	WordE184StyleHeading4 = '\002G\377\373',
	WordE184StyleHeading5 = '\002G\377\372',
	WordE184StyleHeading6 = '\002G\377\371',
	WordE184StyleHeading7 = '\002G\377\370',
	WordE184StyleHeading8 = '\002G\377\367',
	WordE184StyleHeading9 = '\002G\377\366',
	WordE184StyleIndex1 = '\002G\377\365',
	WordE184StyleIndex2 = '\002G\377\364',
	WordE184StyleIndex3 = '\002G\377\363',
	WordE184StyleIndex4 = '\002G\377\362',
	WordE184StyleIndex5 = '\002G\377\361',
	WordE184StyleIndex6 = '\002G\377\360',
	WordE184StyleIndex7 = '\002G\377\357',
	WordE184StyleIndex8 = '\002G\377\356',
	WordE184StyleIndex9 = '\002G\377\355',
	WordE184StyleToc1 = '\002G\377\354',
	WordE184StyleToc2 = '\002G\377\353',
	WordE184StyleToc3 = '\002G\377\352',
	WordE184StyleToc4 = '\002G\377\351',
	WordE184StyleToc5 = '\002G\377\350',
	WordE184StyleToc6 = '\002G\377\347',
	WordE184StyleToc7 = '\002G\377\346',
	WordE184StyleToc8 = '\002G\377\345',
	WordE184StyleToc9 = '\002G\377\344',
	WordE184StyleNormalIndent = '\002G\377\343',
	WordE184StyleFootnoteText = '\002G\377\342',
	WordE184StyleCommentText = '\002G\377\341',
	WordE184StyleHeader = '\002G\377\340',
	WordE184StyleFooter = '\002G\377\337',
	WordE184StyleIndexHeading = '\002G\377\336',
	WordE184StyleCaption = '\002G\377\335',
	WordE184StyleTableOfFigures = '\002G\377\334',
	WordE184StyleFootnoteReference = '\002G\377\331',
	WordE184StyleCommentReference = '\002G\377\330',
	WordE184StyleLineNumber = '\002G\377\327',
	WordE184StylePageNumber = '\002G\377\326',
	WordE184StyleEndnoteReference = '\002G\377\325',
	WordE184StyleEndnoteText = '\002G\377\324',
	WordE184StyleTableOfAuthorities = '\002G\377\323',
	WordE184StyleMacroText = '\002G\377\322',
	WordE184StyleToaHeading = '\002G\377\321',
	WordE184StyleList = '\002G\377\320',
	WordE184StyleListBullet = '\002G\377\317',
	WordE184StyleListNumber = '\002G\377\316',
	WordE184StyleList2 = '\002G\377\315',
	WordE184StyleList3 = '\002G\377\314',
	WordE184StyleList4 = '\002G\377\313',
	WordE184StyleList5 = '\002G\377\312',
	WordE184StyleListBullet2 = '\002G\377\311',
	WordE184StyleListBullet3 = '\002G\377\310',
	WordE184StyleListBullet4 = '\002G\377\307',
	WordE184StyleListBullet5 = '\002G\377\306',
	WordE184StyleListNumber2 = '\002G\377\305',
	WordE184StyleListNumber3 = '\002G\377\304',
	WordE184StyleListNumber4 = '\002G\377\303',
	WordE184StyleListNumber5 = '\002G\377\302',
	WordE184StyleTitle = '\002G\377\301',
	WordE184StyleClosing = '\002G\377\300',
	WordE184StyleSignature = '\002G\377\277',
	WordE184StyleDefaultParagraphFont = '\002G\377\276',
	WordE184StyleBodyTextIndent = '\002G\377\274',
	WordE184StyleListContinue = '\002G\377\273',
	WordE184StyleListContinue2 = '\002G\377\272',
	WordE184StyleListContinue3 = '\002G\377\271',
	WordE184StyleListContinue4 = '\002G\377\270',
	WordE184StyleListContinue5 = '\002G\377\267',
	WordE184StyleMessageHeader = '\002G\377\266',
	WordE184StyleSubtitle = '\002G\377\265',
	WordE184StyleSalutation = '\002G\377\264',
	WordE184StyleDate = '\002G\377\263',
	WordE184StyleBodyTextFirstIndent = '\002G\377\262',
	WordE184StyleBodyTextFirstIndent2 = '\002G\377\261',
	WordE184StyleNoteHeading = '\002G\377\260',
	WordE184StyleBodyText2 = '\002G\377\257',
	WordE184StyleBodyText3 = '\002G\377\256',
	WordE184StyleBodyTextIndent2 = '\002G\377\255',
	WordE184StyleBodyTextIndent3 = '\002G\377\254',
	WordE184StyleBlockQuotation = '\002G\377\253',
	WordE184StyleHyperlink = '\002G\377\252',
	WordE184StyleHyperlinkFollowed = '\002G\377\251',
	WordE184StyleStrong = '\002G\377\250',
	WordE184StyleEmphasis = '\002G\377\247',
	WordE184StyleNavPane = '\002G\377\246',
	WordE184StylePlainText = '\002G\377\245',
	WordE184StyleHtmlNormal = '\002G\377\241',
	WordE184StyleHtmlAcronym = '\002G\377\240',
	WordE184StyleHtmlAddress = '\002G\377\237',
	WordE184StyleHtmlCite = '\002G\377\236',
	WordE184StyleHtmlCode = '\002G\377\235',
	WordE184StyleHtmlDefine = '\002G\377\234',
	WordE184StyleHtmlKeyboard = '\002G\377\233',
	WordE184StyleHtmlPreformatted = '\002G\377\232',
	WordE184StyleHtmlSample = '\002G\377\231',
	WordE184StyleHtmlTypewriter = '\002G\377\230',
	WordE184StyleHtmlVariable = '\002G\377\227',
	WordE184StyleNoteLevel1 = '\002G\377c',
	WordE184StyleNoteLevel2 = '\002G\377b',
	WordE184StyleNoteLevel3 = '\002G\377a',
	WordE184StyleNoteLevel4 = '\002G\377`',
	WordE184StyleNoteLevel5 = '\002G\377_',
	WordE184StyleNoteLevel6 = '\002G\377^',
	WordE184StyleNoteLevel7 = '\002G\377]',
	WordE184StyleNoteLevel8 = '\002G\377\\',
	WordE184StyleNoteLevel9 = '\002G\377[',
	WordE184StyleBibliography = '\002G\377Z',
	WordE184StyleListParagraph = '\002G\377Y',
	WordE184StylePlaceholderText = '\002G\377X'
};
typedef enum WordE184 WordE184;

enum WordE185 {
	WordE185DialogFileDocumentLayoutTabMargins = '\002I\000\017',
	WordE185DialogFileDocumentLayoutTabLayout = '\002I\000\020',
	WordE185DialogFilePageSetupTabMargins = '\002I\000\021',
	WordE185DialogFilePageSetupTabPaperSize = '\002I\000\022',
	WordE185DialogFilePageSetupTabPaperSource = '\002I\000\023',
	WordE185DialogFilePageSetupTabLayout = '\002I\000\024',
	WordE185DialogFilePageSetupTabCharsLines = '\002I\000\025',
	WordE185DialogInsertSymbolTabSymbols = '\002I\000\026',
	WordE185DialogInsertSymbolTabSpecialCharacters = '\002I\000\027',
	WordE185DialogNoteOptionsTabAllFootnotes = '\002I\000\030',
	WordE185DialogNoteOptionsTabAllEndnotes = '\002I\000\031',
	WordE185DialogInsertIndexAndTablesTabIndex = '\002I\000\032',
	WordE185DialogInsertIndexAndTablesTabTableOfContents = '\002I\000\033',
	WordE185DialogInsertIndexAndTablesTabTableOfFigures = '\002I\000\034',
	WordE185DialogInsertIndexAndTablesTabTableOfAuthorities = '\002I\000\035',
	WordE185DialogOrganizerTabStyles = '\002I\000\036',
	WordE185DialogOrganizerTabAutoText = '\002I\000\037',
	WordE185DialogOrganizerTabCommandBars = '\002I\000 ',
	WordE185DialogOrganizerTabMacros = '\002I\000!',
	WordE185DialogFormatFontTabFont = '\002I\000\"',
	WordE185DialogFormatFontTabCharacterSpacing = '\002I\000#',
	WordE185DialogFormatBordersAndShadingTabBorders = '\002I\000%',
	WordE185DialogFormatBordersAndShadingTabPageBorder = '\002I\000&',
	WordE185DialogFormatBordersAndShadingTabShading = '\002I\000\'',
	WordE185DialogToolsEnvelopesAndLabelsTabEnvelopes = '\002I\000(',
	WordE185DialogToolsEnvelopesAndLabelsTabLabels = '\002I\000)',
	WordE185DialogFormatParagraphTabIndentsAndSpacing = '\002I\000*',
	WordE185DialogFormatParagraphTabTextFlow = '\002I\000+',
	WordE185DialogFormatParagraphTabTeisai = '\002I\000,',
	WordE185DialogFormatDrawingObjectTabColorsAndLines = '\002I\000-',
	WordE185DialogFormatDrawingObjectTabSize = '\002I\000.',
	WordE185DialogFormatDrawingObjectTabPosition = '\002I\000/',
	WordE185DialogFormatDrawingObjectTabWrapping = '\002I\0000',
	WordE185DialogFormatDrawingObjectTabPicture = '\002I\0001',
	WordE185DialogFormatDrawingObjectTabTextbox = '\002I\0002',
	WordE185DialogFormatDrawingObjectTabHr = '\002I\0003',
	WordE185DialogToolsAutocorrectExceptionsTabFirstLetter = '\002I\0004',
	WordE185DialogToolsAutocorrectExceptionsTabInitialCaps = '\002I\0005',
	WordE185DialogToolsAutocorrectExceptionsTabHangulAndAlphabet = '\002I\0006',
	WordE185DialogToolsAutocorrectExceptionsTabIac = '\002I\0007',
	WordE185DialogFormatBulletsAndNumberingTabBulleted = '\002I\0008',
	WordE185DialogFormatBulletsAndNumberingTabNumbered = '\002I\0009',
	WordE185DialogFormatBulletsAndNumberingTabOutlineNumbered = '\002I\000:',
	WordE185DialogLetterWizardTabLetterFormat = '\002I\000;',
	WordE185DialogLetterWizardTabRecipientInfo = '\002I\000<',
	WordE185DialogLetterWizardTabOtherElements = '\002I\000=',
	WordE185DialogLetterWizardTabSenderInfo = '\002I\000>',
	WordE185DialogToolsAutoManagerTabAutocorrect = '\002I\000\?',
	WordE185DialogToolsAutoManagerTabMathAutocorrect = '\002I\000@',
	WordE185DialogToolsAutoManagerTabAutoFormatAsYouType = '\002I\000A',
	WordE185DialogToolsAutoManagerTabAutoText = '\002I\000B',
	WordE185DialogToolsAutoManagerTabAutoFormat = '\002I\000C',
	WordE185DialogWebOptionsGeneral = '\002I\000D',
	WordE185DialogWebOptionsFiles = '\002I\000E',
	WordE185DialogWebOptionsPictures = '\002I\000F',
	WordE185DialogWebOptionsEncoding = '\002I\000G',
	WordE185DialogWebOptionsFonts = '\002I\000H',
	WordE185DialogFormatDrawingObjectTabAltText = '\002I\000I'
};
typedef enum WordE185 WordE185;

enum WordE186 {
	WordE186DialogHelpAbout = '\002J\000\011',
	WordE186DialogHelpWordPerfectHelp = '\002J\000\012',
	WordE186DialogHelpWordPerfectHelpOptions = '\002J\001\377',
	WordE186DialogFormatChangeCase = '\002J\001B',
	WordE186DialogToolsOptionsFuzzy = '\002J\003\026',
	WordE186DialogToolsWordCount = '\002J\000\344',
	WordE186DialogDocumentStatistics = '\002J\000N',
	WordE186DialogFileNew = '\002J\000O',
	WordE186DialogFileOpen = '\002J\000P',
	WordE186DialogDataMergeOpenDataSource = '\002J\000Q',
	WordE186DialogDataMergeOpenHeaderSource = '\002J\000R',
	WordE186DialogDataMergeUseAddressBook = '\002J\003\013',
	WordE186DialogFileSaveAs = '\002J\000T',
	WordE186DialogFileSummaryInfo = '\002J\000V',
	WordE186DialogToolsTemplates = '\002J\000W',
	WordE186DialogOrganizer = '\002J\000\336',
	WordE186DialogFilePrint = '\002J\000X',
	WordE186DialogDataMerge = '\002J\002\244',
	WordE186DialogDataMergeCheck = '\002J\002\245',
	WordE186DialogDataMergeQueryOptions = '\002J\002\251',
	WordE186DialogDataMergeFindRecord = '\002J\0029',
	WordE186DialogDataMergeInsertIf = '\002J\017\321',
	WordE186DialogDataMergeInsertNextIf = '\002J\017\325',
	WordE186DialogDataMergeInsertSkipIf = '\002J\017\327',
	WordE186DialogDataMergeInsertFillIn = '\002J\017\320',
	WordE186DialogDataMergeInsertAsk = '\002J\017\317',
	WordE186DialogDataMergeInsertSet = '\002J\017\326',
	WordE186DialogDataMergeHelper = '\002J\002\250',
	WordE186DialogLetterWizard = '\002J\0035',
	WordE186DialogFilePrintSetup = '\002J\000a',
	WordE186DialogFileFind = '\002J\000c',
	WordE186DialogDataMergeCreateDataSource = '\002J\002\202',
	WordE186DialogDataMergeCreateHeaderSource = '\002J\002\203',
	WordE186DialogEditPasteSpecial = '\002J\000o',
	WordE186DialogEditFind = '\002J\000p',
	WordE186DialogEditReplace = '\002J\000u',
	WordE186DialogEditGoToOld = '\002J\003+',
	WordE186DialogEditGoTo = '\002J\003\200',
	WordE186DialogCreateAutoText = '\002J\003h',
	WordE186DialogEditAutoText = '\002J\003\331',
	WordE186DialogEditLinks = '\002J\000|',
	WordE186DialogEditObject = '\002J\000}',
	WordE186DialogConvertObject = '\002J\001\210',
	WordE186DialogTableToText = '\002J\000\200',
	WordE186DialogTextToTable = '\002J\000\177',
	WordE186DialogTableInsertTable = '\002J\000\201',
	WordE186DialogTableInsertCells = '\002J\000\202',
	WordE186DialogTableInsertRow = '\002J\000\203',
	WordE186DialogTableDeleteCells = '\002J\000\205',
	WordE186DialogTableSplitCells = '\002J\000\211',
	WordE186DialogTableFormula = '\002J\001\\',
	WordE186DialogTableAutoFormat = '\002J\0023',
	WordE186DialogTableFormatCell = '\002J\002d',
	WordE186DialogViewZoom = '\002J\002A',
	WordE186DialogNewToolbar = '\002J\002J',
	WordE186DialogInsertBreak = '\002J\000\237',
	WordE186DialogInsertFootnote = '\002J\001r',
	WordE186DialogInsertSymbol = '\002J\000\242',
	WordE186DialogInsertPicture = '\002J\000\243',
	WordE186DialogInsertFile = '\002J\000\244',
	WordE186DialogInsertDateTime = '\002J\000\245',
	WordE186DialogInsertNumber = '\002J\003,',
	WordE186DialogInsertField = '\002J\000\246',
	WordE186DialogInsertDatabase = '\002J\001U',
	WordE186DialogInsertMergeField = '\002J\000\247',
	WordE186DialogInsertBookmark = '\002J\000\250',
	WordE186DialogMarkIndexEntry = '\002J\000\251',
	WordE186DialogMarkCitation = '\002J\001\317',
	WordE186DialogEditToacategory = '\002J\002q',
	WordE186DialogInsertIndexAndTables = '\002J\001\331',
	WordE186DialogInsertIndex = '\002J\000\252',
	WordE186DialogInsertTableOfContents = '\002J\000\253',
	WordE186DialogMarkTableOfContentsEntry = '\002J\001\272',
	WordE186DialogInsertTableOfFigures = '\002J\001\330',
	WordE186DialogInsertTableOfAuthorities = '\002J\001\327',
	WordE186DialogInsertObject = '\002J\000\254',
	WordE186DialogFormatCallout = '\002J\002b',
	WordE186DialogDrawSnapToGrid = '\002J\002y',
	WordE186DialogDrawAlign = '\002J\002z',
	WordE186DialogToolsEnvelopesAndLabels = '\002J\002_',
	WordE186DialogToolsCreateEnvelope = '\002J\000\255',
	WordE186DialogToolsCreateLabels = '\002J\001\351',
	WordE186DialogToolsProtectDocument = '\002J\001\367',
	WordE186DialogToolsProtectSection = '\002J\002B',
	WordE186DialogToolsUnprotectDocument = '\002J\002\011',
	WordE186DialogFormatFont = '\002J\000\256',
	WordE186DialogFormatParagraph = '\002J\000\257',
	WordE186DialogFormatSectionLayout = '\002J\000\260',
	WordE186DialogFormatColumns = '\002J\000\261',
	WordE186DialogFileDocumentLayout = '\002J\000\262',
	WordE186DialogFileMacPageSetup = '\002J\002\255',
	WordE186DialogFilePageSetup = '\002J\000\262',
	WordE186DialogFormatTabs = '\002J\000\263',
	WordE186DialogFormatStyle = '\002J\000\264',
	WordE186DialogFormatStyleGallery = '\002J\001\371',
	WordE186DialogFormatDefineStyleFont = '\002J\000\265',
	WordE186DialogFormatDefineStylePara = '\002J\000\266',
	WordE186DialogFormatDefineStyleTabs = '\002J\000\267',
	WordE186DialogFormatDefineStyleFrame = '\002J\000\270',
	WordE186DialogFormatDefineStyleBorders = '\002J\000\271',
	WordE186DialogFormatDefineStyleLang = '\002J\000\272',
	WordE186DialogFormatPicture = '\002J\000\273',
	WordE186DialogToolsLanguage = '\002J\000\274',
	WordE186DialogFormatBordersAndShading = '\002J\000\275',
	WordE186DialogFormatDrawingObject = '\002J\003\300',
	WordE186DialogFormatFrame = '\002J\000\276',
	WordE186DialogFormatDropCap = '\002J\001\350',
	WordE186DialogFormatBulletsAndNumbering = '\002J\0038',
	WordE186DialogToolsHyphenation = '\002J\000\303',
	WordE186DialogToolsBulletsNumbers = '\002J\000\304',
	WordE186DialogToolsHighlightChanges = '\002J\000\305',
	WordE186DialogToolsAcceptRejectChanges = '\002J\001\372',
	WordE186DialogToolsMergeDocuments = '\002J\001\263',
	WordE186DialogToolsCompareDocuments = '\002J\000\306',
	WordE186DialogTableSort = '\002J\000\307',
	WordE186DialogToolsCustomizeMenuBar = '\002J\002g',
	WordE186DialogToolsCustomize = '\002J\000\230',
	WordE186DialogToolsCustomizeKeyboard = '\002J\001\260',
	WordE186DialogToolsCustomizeMenus = '\002J\001\261',
	WordE186DialogListCommands = '\002J\002\323',
	WordE186DialogToolsOptions = '\002J\003\316',
	WordE186DialogToolsOptionsGeneral = '\002J\000\313',
	WordE186DialogToolsAdvancedSettings = '\002J\000\316',
	WordE186DialogToolsOptionsCompatibility = '\002J\002\015',
	WordE186DialogToolsOptionsPrint = '\002J\000\320',
	WordE186DialogToolsOptionsSave = '\002J\000\321',
	WordE186DialogToolsOptionsSpellingAndGrammar = '\002J\000\323',
	WordE186DialogToolsSpellingAndGrammar = '\002J\003<',
	WordE186DialogToolsThesaurus = '\002J\000\302',
	WordE186DialogToolsOptionsUserInfo = '\002J\000\325',
	WordE186DialogToolsOptionsAutoFormat = '\002J\003\277',
	WordE186DialogToolsOptionsTrackChanges = '\002J\001\202',
	WordE186DialogToolsOptionsEdit = '\002J\000\340',
	WordE186DialogInsertPageNumbers = '\002J\001&',
	WordE186DialogFormatPageNumber = '\002J\001*',
	WordE186DialogNoteOptions = '\002J\001u',
	WordE186DialogCopyFile = '\002J\001,',
	WordE186DialogFormatAddrFonts = '\002J\000g',
	WordE186DialogFormatRetAddrFonts = '\002J\000\335',
	WordE186DialogToolsOptionsFileLocations = '\002J\000\341',
	WordE186DialogToolsCreateDirectory = '\002J\003A',
	WordE186DialogUpdateToc = '\002J\001K',
	WordE186DialogInsertFormField = '\002J\001\343',
	WordE186DialogFormFieldOptions = '\002J\001a',
	WordE186DialogInsertCaption = '\002J\001e',
	WordE186DialogInsertAutoCaption = '\002J\001g',
	WordE186DialogInsertAddCaption = '\002J\001\222',
	WordE186DialogInsertCaptionNumbering = '\002J\001f',
	WordE186DialogInsertCrossReference = '\002J\001o',
	WordE186DialogToolsManageFields = '\002J\002w',
	WordE186DialogToolsAutoManager = '\002J\003\223',
	WordE186DialogToolsAutocorrect = '\002J\001z',
	WordE186DialogToolsAutocorrectExceptions = '\002J\002\372',
	WordE186DialogConnect = '\002J\001\244',
	WordE186DialogToolsOptionsView = '\002J\000\314',
	WordE186DialogInsertSubdocument = '\002J\002G',
	WordE186DialogFileRoutingSlip = '\002J\002p',
	WordE186DialogFontSubstitution = '\002J\002E',
	WordE186DialogToolsOptionsTypography = '\002J\002\343',
	WordE186DialogToolsOptionsAutoFormatAsYouType = '\002J\003\012',
	WordE186DialogControlRun = '\002J\000\353',
	WordE186DialogFileVersions = '\002J\003\261',
	WordE186DialogFileSaveVersion = '\002J\003\357',
	WordE186DialogWindowActivate = '\002J\000\334',
	WordE186DialogToolsMacroRecord = '\002J\000\326',
	WordE186DialogToolsRevisions = '\002J\000\305',
	WordE186DialogWebOptions = '\002J\003\202',
	WordE186DialogFitText = '\002J\003\327',
	WordE186DialogFormatEncloseCharacters = '\002J\004\212',
	WordE186DialogEmail = '\002J\020%',
	WordE186DialogFormatTheme = '\002J\003W',
	WordE186DialogToolsOptionsSecurity = '\002J\005Q',
	WordE186DialogToolsOptionsFeedback = '\002J\006\301',
	WordE186DialogToolsOptionsEditCopyPaste = '\002J\005L',
	WordE186DialogToolsOptionsNoteRecording = '\002J\006a',
	WordE186DialogMathRecognizedFunctions = '\002J\006\321',
	WordE186DialogMathMatrixSpacing = '\002J\006\322',
	WordE186DialogMathEquationArraySpacing = '\002J\006\323'
};
typedef enum WordE186 WordE186;

enum WordE187 {
	WordE187FieldKindNone = '\002K\000\000',
	WordE187FieldKindHot = '\002K\000\001',
	WordE187FieldKindWarm = '\002K\000\002',
	WordE187FieldKindCold = '\002K\000\003'
};
typedef enum WordE187 WordE187;

enum WordE188 {
	WordE188RegularText = '\002L\000\000',
	WordE188NumberText = '\002L\000\001',
	WordE188DateText = '\002L\000\002',
	WordE188CurrentDateText = '\002L\000\003',
	WordE188CurrentTimeText = '\002L\000\004',
	WordE188CalculationText = '\002L\000\005'
};
typedef enum WordE188 WordE188;

enum WordE189 {
	WordE189NeverConvert = '\002M\000\000',
	WordE189AlwaysConvert = '\002M\000\001',
	WordE189AskToNotConvert = '\002M\000\002',
	WordE189AskToConvert = '\002M\000\003'
};
typedef enum WordE189 WordE189;

enum WordE190 {
	WordE190NotAMergeDocument = '\002M\377\377',
	WordE190DocumentTypeFormLetters = '\002N\000\000',
	WordE190DocumentTypeMailingLabels = '\002N\000\001',
	WordE190DocumentTypeEnvelopes = '\002N\000\002',
	WordE190DocumentTypeCatalog = '\002N\000\003'
};
typedef enum WordE190 WordE190;

enum WordE191 {
	WordE191NormalDocument = '\002O\000\000',
	WordE191MainDocumentOnly = '\002O\000\001',
	WordE191MainAndDataSource = '\002O\000\002',
	WordE191MainAndHeader = '\002O\000\003',
	WordE191MainAndSourceAndHeader = '\002O\000\004',
	WordE191DataSource = '\002O\000\005'
};
typedef enum WordE191 WordE191;

enum WordE192 {
	WordE192SendToNewDocument = '\002P\000\000',
	WordE192SendToPrinter = '\002P\000\001',
	WordE192SendToEmail = '\002P\000\002',
	WordE192SendToFax = '\002P\000\003'
};
typedef enum WordE192 WordE192;

enum WordE193 {
	WordE193NoActiveRecord = '\002P\377\377',
	WordE193NextDataRecord = '\002P\377\376',
	WordE193PreviousDataRecord = '\002P\377\375',
	WordE193FirstDataRecord = '\002P\377\374',
	WordE193LastDataRecord = '\002P\377\373'
};
typedef enum WordE193 WordE193;

enum WordE194 {
	WordE194DefaultFirstRecord = '\000\000\000\001',
	WordE194DefaultLastRecord = '\377\377\377\360'
};
typedef enum WordE194 WordE194;

enum WordE195 {
	WordE195NoMergeInfo = '\002R\377\377',
	WordE195MergeInfoFromWord = '\002S\000\000',
	WordE195MergeInfoFromAccessDde = '\002S\000\001',
	WordE195MergeInfoFromExcelDde = '\002S\000\002',
	WordE195MergeInfoFromMsqueryDde = '\002S\000\003',
	WordE195MergeInfoFromOdbc = '\002S\000\004'
};
typedef enum WordE195 WordE195;

enum WordE196 {
	WordE196MergeIfEqual = '\002T\000\000',
	WordE196MergeIfNotEqual = '\002T\000\001',
	WordE196MergeIfLessThan = '\002T\000\002',
	WordE196MergeIfGreaterThan = '\002T\000\003',
	WordE196MergeIfLessThanOrEqual = '\002T\000\004',
	WordE196MergeIfGreaterThanOrEqual = '\002T\000\005',
	WordE196MergeIfIsBlank = '\002T\000\006',
	WordE196MergeIfIsNotBlank = '\002T\000\007'
};
typedef enum WordE196 WordE196;

enum WordE197 {
	WordE197SortByName = '\002U\000\000',
	WordE197SortByLocation = '\002U\000\001'
};
typedef enum WordE197 WordE197;

enum WordE198 {
	WordE198WindowStateNormal = '\002V\000\000',
	WordE198WindowStateMaximize = '\002V\000\001',
	WordE198WindowStateMinimize = '\002V\000\002'
};
typedef enum WordE198 WordE198;

enum WordE199 {
	WordE199LinkNone = '\002W\000\000',
	WordE199LinkDataInDoc = '\002W\000\001',
	WordE199LinkDataOnDisk = '\002W\000\002'
};
typedef enum WordE199 WordE199;

enum WordE200 {
	WordE200LinkTypeOle = '\002X\000\000',
	WordE200LinkTypePicture = '\002X\000\001',
	WordE200LinkTypeText = '\002X\000\002',
	WordE200LinkTypeReference = '\002X\000\003',
	WordE200LinkTypeInclude = '\002X\000\004',
	WordE200LinkTypeImport = '\002X\000\005',
	WordE200LinkTypeDde = '\002X\000\006',
	WordE200LinkTypeDdeauto = '\002X\000\007',
	WordE200LinkTypeChart = '\002X\000\010'
};
typedef enum WordE200 WordE200;

enum WordE201 {
	WordE201WindowDocument = '\002Y\000\000',
	WordE201WindowTemplate = '\002Y\000\001'
};
typedef enum WordE201 WordE201;

enum WordE202 {
	WordE202NormalView = '\002Z\000\001',
	WordE202DraftView = '\002Z\000\001',
	WordE202OutlineView = '\002Z\000\002',
	WordE202PageView = '\002Z\000\003',
	WordE202PrintView = '\002Z\000\003',
	WordE202PrintPreviewView = '\002Z\000\004',
	WordE202MasterView = '\002Z\000\005',
	WordE202OnlineView = '\002Z\000\006',
	WordE202WordNoteView = '\002Z\000\007',
	WordE202PublishingView = '\002Z\000\010',
	WordE202ConflictView = '\002Z\000\011'
};
typedef enum WordE202 WordE202;

enum WordE203 {
	WordE203SeekMainDocument = '\002[\000\000',
	WordE203SeekPrimaryHeader = '\002[\000\001',
	WordE203SeekFirstPageHeader = '\002[\000\002',
	WordE203SeekEvenPagesHeader = '\002[\000\003',
	WordE203SeekPrimaryFooter = '\002[\000\004',
	WordE203SeekFirstPageFooter = '\002[\000\005',
	WordE203SeekEvenPagesFooter = '\002[\000\006',
	WordE203SeekFootnotes = '\002[\000\007',
	WordE203SeekEndnotes = '\002[\000\010',
	WordE203SeekCurrentPageHeader = '\002[\000\011',
	WordE203SeekCurrentPageFooter = '\002[\000\012'
};
typedef enum WordE203 WordE203;

enum WordE204 {
	WordE204PaneNone = '\002\\\000\000',
	WordE204PanePrimaryHeader = '\002\\\000\001',
	WordE204PaneFirstPageHeader = '\002\\\000\002',
	WordE204PaneEvenPagesHeader = '\002\\\000\003',
	WordE204PanePrimaryFooter = '\002\\\000\004',
	WordE204PaneFirstPageFooter = '\002\\\000\005',
	WordE204PaneEvenPagesFooter = '\002\\\000\006',
	WordE204PaneFootnotes = '\002\\\000\007',
	WordE204PaneEndnotes = '\002\\\000\010',
	WordE204PaneFootnoteContinuationNotice = '\002\\\000\011',
	WordE204PaneFootnoteContinuationSeparator = '\002\\\000\012',
	WordE204PaneFootnoteSeparator = '\002\\\000\013',
	WordE204PaneEndnoteContinuationNotice = '\002\\\000\014',
	WordE204PaneEndnoteContinuationSeparator = '\002\\\000\015',
	WordE204PaneEndnoteSeparator = '\002\\\000\016',
	WordE204PaneComments = '\002\\\000\017',
	WordE204PaneCurrentPageHeader = '\002\\\000\020',
	WordE204PaneCurrentPageFooter = '\002\\\000\021',
	WordE204PaneRevisions = '\002\\\000\022'
};
typedef enum WordE204 WordE204;

enum WordE205 {
	WordE205PageFitNone = '\002]\000\000',
	WordE205PageFitFullPage = '\002]\000\001',
	WordE205PageFitBestFit = '\002]\000\002'
};
typedef enum WordE205 WordE205;

enum WordE206 {
	WordE206BrowsePage = '\002^\000\001',
	WordE206BrowseSection = '\002^\000\002',
	WordE206BrowseComment = '\002^\000\003',
	WordE206BrowseFootnote = '\002^\000\004',
	WordE206BrowseEndnote = '\002^\000\005',
	WordE206BrowseField = '\002^\000\006',
	WordE206BrowseTable = '\002^\000\007',
	WordE206BrowseGraphic = '\002^\000\010',
	WordE206BrowseHeading = '\002^\000\011',
	WordE206BrowseEdit = '\002^\000\012',
	WordE206BrowseFind = '\002^\000\013',
	WordE206BrowseGoTo = '\002^\000\014'
};
typedef enum WordE206 WordE206;

enum WordE207 {
	WordE207PrinterDefaultBin = '\002_\000\000',
	WordE207PrinterUpperBin = '\002_\000\001',
	WordE207PrinterOnlyBin = '\002_\000\001',
	WordE207PrinterLowerBin = '\002_\000\002',
	WordE207PrinterMiddleBin = '\002_\000\003',
	WordE207PrinterManualFeed = '\002_\000\004',
	WordE207PrinterEnvelopeFeed = '\002_\000\005',
	WordE207PrinterManualEnvelopeFeed = '\002_\000\006',
	WordE207PrinterAutomaticSheetFeed = '\002_\000\007',
	WordE207PrinterTractorFeed = '\002_\000\010',
	WordE207PrinterSmallFormatBin = '\002_\000\011',
	WordE207PrinterLargeFormatBin = '\002_\000\012',
	WordE207PrinterLargeCapacityBin = '\002_\000\013',
	WordE207PrinterPaperCassette = '\002_\000\016',
	WordE207PrinterFormSource = '\002_\000\017'
};
typedef enum WordE207 WordE207;

enum WordE208 {
	WordE208OrientPortrait = '\002`\000\000',
	WordE208OrientLandscape = '\002`\000\001'
};
typedef enum WordE208 WordE208;

enum WordE209 {
	WordE209NoSelection = '\002a\000\000',
	WordE209SelectionIp = '\002a\000\001',
	WordE209SelectionNormal = '\002a\000\002',
	WordE209SelectionFrame = '\002a\000\003',
	WordE209SelectionColumn = '\002a\000\004',
	WordE209SelectionRow = '\002a\000\005',
	WordE209SelectionBlock = '\002a\000\006',
	WordE209SelectionInlineShape = '\002a\000\007',
	WordE209SelectionShape = '\002a\000\010'
};
typedef enum WordE209 WordE209;

enum WordE210 {
	WordE210CaptionFigure = '\002a\377\377',
	WordE210CaptionTable = '\002a\377\376',
	WordE210CaptionEquation = '\002a\377\375'
};
typedef enum WordE210 WordE210;

enum WordE211 {
	WordE211ReferenceTypeNumberedItem = '\002c\000\000',
	WordE211ReferenceTypeHeading = '\002c\000\001',
	WordE211ReferenceTypeBookmark = '\002c\000\002',
	WordE211ReferenceTypeFootnote = '\002c\000\003',
	WordE211ReferenceTypeEndnote = '\002c\000\004'
};
typedef enum WordE211 WordE211;

enum WordE212 {
	WordE212ReferenceContentText = '\002c\377\377',
	WordE212ReferenceNumberRelativeContext = '\002c\377\376',
	WordE212ReferenceNumberNoContext = '\002c\377\375',
	WordE212ReferenceNumberFullContext = '\002c\377\374',
	WordE212ReferenceEntireCaption = '\002d\000\002',
	WordE212ReferenceOnlyLabelAndNumber = '\002d\000\003',
	WordE212ReferenceOnlyCaptionText = '\002d\000\004',
	WordE212ReferenceFootnoteNumber = '\002d\000\005',
	WordE212ReferenceEndnoteNumber = '\002d\000\006',
	WordE212ReferencePageNumber = '\002d\000\007',
	WordE212ReferencePosition = '\002d\000\017',
	WordE212ReferenceFootnoteNumberFormatted = '\002d\000\020',
	WordE212ReferenceEndnoteNumberFormatted = '\002d\000\021'
};
typedef enum WordE212 WordE212;

enum WordE213 {
	WordE213IndexTemplate = '\002e\000\000',
	WordE213IndexClassic = '\002e\000\001',
	WordE213IndexFancy = '\002e\000\002',
	WordE213IndexModern = '\002e\000\003',
	WordE213IndexBulleted = '\002e\000\004',
	WordE213IndexFormal = '\002e\000\005',
	WordE213IndexSimple = '\002e\000\006'
};
typedef enum WordE213 WordE213;

enum WordE214 {
	WordE214IndexIndent = '\002f\000\000',
	WordE214IndexRunin = '\002f\000\001'
};
typedef enum WordE214 WordE214;

enum WordE215 {
	WordE215WrapNever = '\002g\000\000',
	WordE215WrapAlways = '\002g\000\001',
	WordE215WrapAsk = '\002g\000\002'
};
typedef enum WordE215 WordE215;

enum WordE216 {
	WordE216NoRevision = '\002h\000\000',
	WordE216RevisionInsert = '\002h\000\001',
	WordE216RevisionDelete = '\002h\000\002',
	WordE216RevisionProperty = '\002h\000\003',
	WordE216RevisionParagraphNumber = '\002h\000\004',
	WordE216RevisionDisplayField = '\002h\000\005',
	WordE216RevisionReconcile = '\002h\000\006',
	WordE216RevisionConflict = '\002h\000\007',
	WordE216RevisionStyle = '\002h\000\010',
	WordE216RevisionReplace = '\002h\000\011',
	WordE216RevisionParagraphProperty = '\002h\000\012',
	WordE216RevisionTableProperty = '\002h\000\013',
	WordE216RevisionSectionProperty = '\002h\000\014',
	WordE216RevisionStyleDefinition = '\002h\000\015',
	WordE216RevisionMoveFrom = '\002h\000\016',
	WordE216RevisionMoveTo = '\002h\000\017',
	WordE216RevisionCellInsertion = '\002h\000\020',
	WordE216RevisionCellDeletion = '\002h\000\021',
	WordE216RevisionCellMerge = '\002h\000\022',
	WordE216RevisionCellSplit = '\002h\000\023',
	WordE216RevisionConflictInsert = '\002h\000\024',
	WordE216RevisionConflictDelete = '\002h\000\025'
};
typedef enum WordE216 WordE216;

enum WordE217 {
	WordE217OneAfterAnother = '\002i\000\000',
	WordE217AllAtOnce = '\002i\000\001'
};
typedef enum WordE217 WordE217;

enum WordE218 {
	WordE218NotYetRouted = '\002j\000\000',
	WordE218RouteInProgress = '\002j\000\001',
	WordE218RouteComplete = '\002j\000\002'
};
typedef enum WordE218 WordE218;

enum WordE219 {
	WordE219SectionContinuous = '\002k\000\000',
	WordE219SectionNewColumn = '\002k\000\001',
	WordE219SectionNewPage = '\002k\000\002',
	WordE219SectionEvenPage = '\002k\000\003',
	WordE219SectionOddPage = '\002k\000\004'
};
typedef enum WordE219 WordE219;

enum WordE220 {
	WordE220DoNotSaveChanges = '\002l\000\000',
	WordE220SaveChanges = '\002k\377\377',
	WordE220PromptToSaveChanges = '\002k\377\376'
};
typedef enum WordE220 WordE220;

enum WordE221 {
	WordE221DocumentNotSpecified = '\002m\000\000',
	WordE221DocumentLetter = '\002m\000\001',
	WordE221DocumentEmail = '\002m\000\002'
};
typedef enum WordE221 WordE221;

enum WordE222 {
	WordE222TypeDocument = '\002n\000\000',
	WordE222TypeTemplate = '\002n\000\001'
};
typedef enum WordE222 WordE222;

enum WordE223 {
	WordE223WordDocument = '\002o\000\000',
	WordE223OriginalDocumentFormat = '\002o\000\001',
	WordE223PromptUser = '\002o\000\002'
};
typedef enum WordE223 WordE223;

enum WordE224 {
	WordE224RelocateUp = '\002p\000\000',
	WordE224RelocateDown = '\002p\000\001'
};
typedef enum WordE224 WordE224;

enum WordE225 {
	WordE225InsertedTextMarkNone = '\002q\000\000',
	WordE225InsertedTextMarkBold = '\002q\000\001',
	WordE225InsertedTextMarkItalic = '\002q\000\002',
	WordE225InsertedTextMarkUnderline = '\002q\000\003',
	WordE225InsertedTextMarkDoubleUnderline = '\002q\000\004',
	WordE225InsertedTextMarkColorOnly = '\002q\000\005',
	WordE225InsertedTextMarkStrikeThrough = '\002q\000\006',
	WordE225InsertedTextMarkDoubleStrikeThrough = '\002q\000\007'
};
typedef enum WordE225 WordE225;

enum WordE226 {
	WordE226RevisedLinesMarkNone = '\002r\000\000',
	WordE226RevisedLinesMarkLeftBorder = '\002r\000\001',
	WordE226RevisedLinesMarkRightBorder = '\002r\000\002',
	WordE226RevisedLinesMarkOutsideBorder = '\002r\000\003'
};
typedef enum WordE226 WordE226;

enum WordE227 {
	WordE227DeletedTextMarkHidden = '\002s\000\000',
	WordE227DeletedTextMarkStrikeThrough = '\002s\000\001',
	WordE227DeletedTextMarkCaret = '\002s\000\002',
	WordE227DeletedTextMarkPound = '\002s\000\003',
	WordE227DeletedTextMarkNone = '\002s\000\004',
	WordE227DeletedTextMarkBold = '\002s\000\005',
	WordE227DeletedTextMarkItalic = '\002s\000\006',
	WordE227DeletedTextMarkUnderline = '\002s\000\007',
	WordE227DeletedTextMarkDoubleUnderline = '\002s\000\010',
	WordE227DeletedTextMarkColorOnly = '\002s\000\011',
	WordE227DeletedTextMarkDoubleStrikeThrough = '\002s\000\012'
};
typedef enum WordE227 WordE227;

enum WordE228 {
	WordE228RevisedPropertiesMarkNone = '\002t\000\000',
	WordE228RevisedPropertiesMarkBold = '\002t\000\001',
	WordE228RevisedPropertiesMarkItalic = '\002t\000\002',
	WordE228RevisedPropertiesMarkUnderline = '\002t\000\003',
	WordE228RevisedPropertiesMarkDoubleUnderline = '\002t\000\004',
	WordE228RevisedPropertiesMarkColorOnly = '\002t\000\005',
	WordE228RevisedPropertiesMarkStrikeThrough = '\002t\000\006',
	WordE228RevisedPropertiesMarkDoubleStrikeThrough = '\002t\000\007'
};
typedef enum WordE228 WordE228;

enum WordE316 {
	WordE316MoveToTextMarkNone = '\002\314\000\000',
	WordE316MoveToTextMarkBold = '\002\314\000\001',
	WordE316MoveToTextMarkItalic = '\002\314\000\002',
	WordE316MoveToTextMarkUnderline = '\002\314\000\003',
	WordE316MoveToTextMarkDoubleUnderline = '\002\314\000\004',
	WordE316MoveToTextMarkColorOnly = '\002\314\000\005',
	WordE316MoveToTextMarkStrikeThrough = '\002\314\000\006',
	WordE316MoveToTextMarkDoubleStrikeThrough = '\002\314\000\007'
};
typedef enum WordE316 WordE316;

enum WordE319 {
	WordE319MathAccentType = '\002\317\000\001',
	WordE319MathFunctionBarType = '\002\317\000\002',
	WordE319MathBoxType = '\002\317\000\003',
	WordE319MathBoxedFormulaType = '\002\317\000\004',
	WordE319MathDelimitersType = '\002\317\000\005',
	WordE319MathEquationArrayType = '\002\317\000\006',
	WordE319MathFractionType = '\002\317\000\007',
	WordE319MathFunctionApplyType = '\002\317\000\010',
	WordE319MathStretchStackType = '\002\317\000\011',
	WordE319MathLowerLimitType = '\002\317\000\012',
	WordE319MathUpperLimitType = '\002\317\000\013',
	WordE319MathMatrixType = '\002\317\000\014',
	WordE319MathNaryOperatorType = '\002\317\000\015',
	WordE319MathPhantomType = '\002\317\000\016',
	WordE319MathLeftSubSuperscriptType = '\002\317\000\017',
	WordE319MathRadicalType = '\002\317\000\020',
	WordE319MathSubscriptType = '\002\317\000\021',
	WordE319MathSubSuperscriptType = '\002\317\000\022',
	WordE319MathSuperscriptType = '\002\317\000\023',
	WordE319MathTextType = '\002\317\000\024',
	WordE319MathNormalTextType = '\002\317\000\025',
	WordE319MathLiteralTextType = '\002\317\000\026'
};
typedef enum WordE319 WordE319;

enum WordE320 {
	WordE320EquationHorizontalAlignCenter = '\002\320\000\000',
	WordE320EquationHorizontalAlignLeft = '\002\320\000\001',
	WordE320EquationHorizontalAlignRight = '\002\320\000\002'
};
typedef enum WordE320 WordE320;

enum WordE321 {
	WordE321EquationVerticalAlignCenter = '\002\321\000\000',
	WordE321EquationVerticalAlignTop = '\002\321\000\001',
	WordE321EquationVerticalAlignBottom = '\002\321\000\002'
};
typedef enum WordE321 WordE321;

enum WordE322 {
	WordE322MathFractionTypeBar = '\002\322\000\000',
	WordE322MathFractionTypeStack = '\002\322\000\001',
	WordE322MathFractionTypeSlashed = '\002\322\000\002',
	WordE322MathFractionTypeLinear = '\002\322\000\003'
};
typedef enum WordE322 WordE322;

enum WordE323 {
	WordE323EquationSpacingSingle = '\002\323\000\000',
	WordE323EquationSpacingOneAndAHalf = '\002\323\000\001',
	WordE323EquationSpacingDouble = '\002\323\000\002',
	WordE323EquationSpacingExactly = '\002\323\000\003',
	WordE323EquationSpacingMultiple = '\002\323\000\004'
};
typedef enum WordE323 WordE323;

enum WordE324 {
	WordE324EquationDisplayProfessional = '\002\324\000\000' /* Specifies an equation is displayed in professional/built up format. */,
	WordE324EquationDisplayInline = '\002\324\000\001' /* Specifies an equation is displayed in linear/built down style. */
};
typedef enum WordE324 WordE324;

enum WordE325 {
	WordE325MathDelimitersCenterVertically = '\002\325\000\000',
	WordE325MathDelimitersMatchContentSize = '\002\325\000\001'
};
typedef enum WordE325 WordE325;

enum WordE326 {
	WordE326EquationJustificationCenterGroup = '\002\326\000\001',
	WordE326EquationJustificationCenter = '\002\326\000\002',
	WordE326EquationJustificationLeft = '\002\326\000\003',
	WordE326EquationJustificationRight = '\002\326\000\004',
	WordE326EquationJustificationInline = '\002\326\000\007'
};
typedef enum WordE326 WordE326;

enum WordE327 {
	WordE327MathBinaryOperatorBreakBefore = '\002\327\000\000',
	WordE327MathBinaryOperatorBreakAfter = '\002\327\000\001',
	WordE327MathBinaryOperatorRepeat = '\002\327\000\002'
};
typedef enum WordE327 WordE327;

enum WordE328 {
	WordE328MinusMinus = '\002\330\000\000',
	WordE328PlusMinus = '\002\330\000\001',
	WordE328MinusPlus = '\002\330\000\002'
};
typedef enum WordE328 WordE328;

enum WordE329 {
	WordE329BuildingBlockEquations = '\000\000\000\003'
};
typedef enum WordE329 WordE329;

enum WordE330 {
	WordE330InlineBuildingBlock = '\000\000\000\000',
	WordE330PageLevelBuildingBlock = '\000\000\000\001',
	WordE330ParagraphLevelBuildingBlock = '\000\000\000\002'
};
typedef enum WordE330 WordE330;

enum WordE317 {
	WordE317MoveFromTextMarkHidden = '\002\315\000\000',
	WordE317MoveFromTextMarkDoubleStrikeThrough = '\002\315\000\001',
	WordE317MoveFromTextMarkStrikeThrough = '\002\315\000\002',
	WordE317MoveFromTextMarkCaret = '\002\315\000\003',
	WordE317MoveFromTextMarkPound = '\002\315\000\004',
	WordE317MoveFromTextMarkNone = '\002\315\000\005',
	WordE317MoveFromTextMarkBold = '\002\315\000\006',
	WordE317MoveFromTextMarkItalic = '\002\315\000\007',
	WordE317MoveFromTextMarkUnderline = '\002\315\000\010',
	WordE317MoveFromTextMarkDoubleUnderline = '\002\315\000\011',
	WordE317MoveFromTextMarkColorOnly = '\002\315\000\012'
};
typedef enum WordE317 WordE317;

enum WordE318 {
	WordE318CellColorByAuthor = '\002\315\377\377',
	WordE318CellColorNoHighlight = '\002\316\000\000',
	WordE318CellColorPink = '\002\316\000\001',
	WordE318CellColorLightBlue = '\002\316\000\002',
	WordE318CellColorLightYellow = '\002\316\000\003',
	WordE318CellColorLightPurple = '\002\316\000\004',
	WordE318CellColorLightOrange = '\002\316\000\005',
	WordE318CellColorLightGreen = '\002\316\000\006',
	WordE318CellColorLightGray = '\002\316\000\007'
};
typedef enum WordE318 WordE318;

enum WordE229 {
	WordE229FieldShadingNever = '\002u\000\000',
	WordE229FieldShadingAlways = '\002u\000\001',
	WordE229FieldShadingWhenSelected = '\002u\000\002'
};
typedef enum WordE229 WordE229;

enum WordE230 {
	WordE230DocumentsPath = '\002v\000\000',
	WordE230PicturesPath = '\002v\000\001',
	WordE230UserTemplatesPath = '\002v\000\002',
	WordE230WorkgroupTemplatesPath = '\002v\000\003',
	WordE230UserOptionsPath = '\002v\000\004',
	WordE230AutoRecoverPath = '\002v\000\005',
	WordE230ToolsPath = '\002v\000\006',
	WordE230TutorialPath = '\002v\000\007',
	WordE230StartupPath = '\002v\000\010',
	WordE230ProgramPath = '\002v\000\011',
	WordE230GraphicsFiltersPath = '\002v\000\012',
	WordE230TextConvertersPath = '\002v\000\013',
	WordE230ProofingToolsPath = '\002v\000\014',
	WordE230TempFilePath = '\002v\000\015',
	WordE230CurrentFolderPath = '\002v\000\016',
	WordE230StyleGalleryPath = '\002v\000\017',
	WordE230TrashPath = '\002v\000\020',
	WordE230OfficePath = '\002v\000\021',
	WordE230TypeLibrariesPath = '\002v\000\022',
	WordE230BorderArtPath = '\002v\000\023'
};
typedef enum WordE230 WordE230;

enum WordE231 {
	WordE231NoTabHangingIndent = '\002w\000\001',
	WordE231NoSpaceForRaisedOrLoweredCharacters = '\002w\000\002',
	WordE231PrintColorsBlack = '\002w\000\003',
	WordE231WrapTrailSpaces = '\002w\000\004',
	WordE231NoColumnBalance = '\002w\000\005',
	WordE231ConvertDataMergeEscapes = '\002w\000\006',
	WordE231SuppressSpaceBeforeAfterPageBreak = '\002w\000\007',
	WordE231SuppressTopSpacing = '\002w\000\010',
	WordE231OriginalWordTableRules = '\002w\000\011',
	WordE231TransparentMetafiles = '\002w\000\012',
	WordE231ShowBreaksInFrames = '\002w\000\013',
	WordE231SwapBordersFacingPages = '\002w\000\014',
	WordE231LeaveBackslashAlone = '\002w\000\015',
	WordE231ExpandShiftReturn = '\002w\000\016',
	WordE231DoNotUnderlineTrailingSpaces = '\002w\000\017',
	WordE231DoNotBalanceSBCSAndDBCSCharacters = '\002w\000\020',
	WordE231SuppressTopSpacingMacWord5 = '\002w\000\021',
	WordE231SpacingInWholePoints = '\002w\000\022',
	WordE231PrintBodyTextBeforeHeader = '\002w\000\023',
	WordE231NoExtraSpaceBetweenRowsOfText = '\002w\000\024',
	WordE231NoSpaceForUnderlines = '\002w\000\025',
	WordE231UseLargerSmallCaps = '\002w\000\026',
	WordE231NoExtraLineSpacing = '\002w\000\027',
	WordE231TruncateFontHeight = '\002w\000\030',
	WordE231SubstituteFontBySize = '\002w\000\031',
	WordE231UsePrinterMetrics = '\002w\000\032',
	WordE231Word6BorderRules = '\002w\000\033',
	WordE231ExactOnTop = '\002w\000\034',
	WordE231SuppressBottomSpacing = '\002w\000\035',
	WordE231WordPerfectSpaceWidth = '\002w\000\036',
	WordE231WordPerfectJustification = '\002w\000\037',
	WordE231Word6LineWrap = '\002w\000 ',
	WordE231Word96ShapeLayout = '\002w\000!',
	WordE231Word98FootnoteLayout = '\002w\000\"',
	WordE231DoNotUseHtmlParagraphAutoSpacing = '\002w\000#',
	WordE231DoNotAdjustLineHeightInTable = '\002w\000$',
	WordE231ForgetLastTabAlignment = '\002w\000%',
	WordE231Word95AutoSpace = '\002w\000&',
	WordE231AlignTablesRowByRow = '\002w\000\'',
	WordE231LayoutRawTableWidth = '\002w\000(',
	WordE231LayoutTableRowsApart = '\002w\000)',
	WordE231UseWord97LineBreakingRules = '\002w\000*',
	WordE231DoNotBreakWrappedTables = '\002w\000+',
	WordE231GrowAutofit = '\002w\000,',
	WordE231DoNotSnapTextToGridInTableWithObjects = '\002w\000-',
	WordE231SelectFieldWithFirstOrLastCharacter = '\002w\000.',
	WordE231ApplyBreakingRules = '\002w\000/',
	WordE231DoNotWrapTextWithPunctuation = '\002w\0000',
	WordE231DoNotUseAsianBreakRulesInGrid = '\002w\0001',
	WordE231UseWord2002TableStyleRules = '\002w\0002'
};
typedef enum WordE231 WordE231;

enum WordE314 {
	WordE314DefaultCompatibilitySettings = '\002\312\000\000' /* Microsoft Word 2007-2008 */,
	WordE314Word20002004AndX = '\002\312\000\014' /* Microsoft Word 2000-2004 and X */,
	WordE314Word97And98 = '\002\312\000\001' /* Microsoft Word 97-98 */,
	WordE314Word6And95 = '\002\312\000\002' /* Microsoft Word 6.0/95 */,
	WordE314WinWord1 = '\002\312\000\003' /* Word for Windows 1.0 */,
	WordE314WinWord2 = '\002\312\000\004' /* Word for Windows 2.0 */,
	WordE314MacWord5 = '\002\312\000\005' /* Word for the Macintosh 5.x */,
	WordE314DosWord = '\002\312\000\006' /* Word for MS-DOS */,
	WordE314Wordperfect5 = '\002\312\000\007' /* WordPerfect 5.x */,
	WordE314WinWordperfect6 = '\002\312\000\010' /* WordPerfect 6.x for Windows */,
	WordE314DosWordperfect6 = '\002\312\000\011' /* WordPerfect 6.0 for DOS */,
	WordE314AsianWord97And98 = '\002\312\000\012' /* Microsoft Word Asian Versions 97-98 */,
	WordE314UsWord6And95 = '\002\312\000\013' /* US Microsoft Word for Windows 6.0/95 */,
	WordE314CustomCompatibilitySettings = '\002\312\000\015' /* Custom settings */
};
typedef enum WordE314 WordE314;

enum WordE232 {
	WordE232PaperTenXFourteen = '\002x\000\000',
	WordE232PaperElevenXSeventeen = '\002x\000\001',
	WordE232PaperLetter = '\002x\000\002',
	WordE232PaperLetterSmall = '\002x\000\003',
	WordE232PaperLegal = '\002x\000\004',
	WordE232PaperExecutive = '\002x\000\005',
	WordE232PaperA3 = '\002x\000\006',
	WordE232PaperA4 = '\002x\000\007',
	WordE232PaperA4Small = '\002x\000\010',
	WordE232PaperA5 = '\002x\000\011',
	WordE232PaperB4 = '\002x\000\012',
	WordE232PaperB5 = '\002x\000\013',
	WordE232PaperCsheet = '\002x\000\014',
	WordE232PaperDsheet = '\002x\000\015',
	WordE232PaperEsheet = '\002x\000\016',
	WordE232PaperFanfoldLegalGerman = '\002x\000\017',
	WordE232PaperFanfoldStdGerman = '\002x\000\020',
	WordE232PaperFanfoldUs = '\002x\000\021',
	WordE232PaperFolio = '\002x\000\022',
	WordE232PaperLedger = '\002x\000\023',
	WordE232PaperNote = '\002x\000\024',
	WordE232PaperQuarto = '\002x\000\025',
	WordE232PaperStatement = '\002x\000\026',
	WordE232PaperTabloid = '\002x\000\027',
	WordE232PaperEnvelope9 = '\002x\000\030',
	WordE232PaperEnvelope10 = '\002x\000\031',
	WordE232PaperEnvelope11 = '\002x\000\032',
	WordE232PaperEnvelope12 = '\002x\000\033',
	WordE232PaperEnvelope14 = '\002x\000\034',
	WordE232PaperEnvelopeB4 = '\002x\000\035',
	WordE232PaperEnvelopeB5 = '\002x\000\036',
	WordE232PaperEnvelopeB6 = '\002x\000\037',
	WordE232PaperEnvelopeC3 = '\002x\000 ',
	WordE232PaperEnvelopeC4 = '\002x\000!',
	WordE232PaperEnvelopeC5 = '\002x\000\"',
	WordE232PaperEnvelopeC6 = '\002x\000#',
	WordE232PaperEnvelopeC65 = '\002x\000$',
	WordE232PaperEnvelopeDl = '\002x\000%',
	WordE232PaperEnvelopeItaly = '\002x\000&',
	WordE232PaperEnvelopeMonarch = '\002x\000\'',
	WordE232PaperEnvelopePersonal = '\002x\000(',
	WordE232PaperCustom = '\002x\000)'
};
typedef enum WordE232 WordE232;

enum WordE233 {
	WordE233CustomLabelLetter = '\002y\000\000',
	WordE233CustomLabelLetterLandscape = '\002y\000\001',
	WordE233CustomLabelA4 = '\002y\000\002',
	WordE233CustomLabelA4Landscape = '\002y\000\003',
	WordE233CustomLabelA5 = '\002y\000\004',
	WordE233CustomLabelA5Landscape = '\002y\000\005',
	WordE233CustomLabelB5 = '\002y\000\006',
	WordE233CustomLabelMini = '\002y\000\007',
	WordE233CustomLabelFanfold = '\002y\000\010'
};
typedef enum WordE233 WordE233;

enum WordE234 {
	WordE234NoDocumentProtection = '\002y\377\377',
	WordE234AllowOnlyRevisions = '\002z\000\000',
	WordE234AllowOnlyComments = '\002z\000\001',
	WordE234AllowOnlyFormFields = '\002z\000\002',
	WordE234AllowOnlyReading = '\002z\000\003'
};
typedef enum WordE234 WordE234;

enum WordE235 {
	WordE235Adjective = '\002{\000\000',
	WordE235Noun = '\002{\000\001',
	WordE235Adverb = '\002{\000\002',
	WordE235Verb = '\002{\000\003',
	WordE235Pronoun = '\002{\000\004',
	WordE235Conjunction = '\002{\000\005',
	WordE235Preposition = '\002{\000\006',
	WordE235Interjection = '\002{\000\007',
	WordE235Idiom = '\002{\000\010',
	WordE235Other = '\002{\000\011'
};
typedef enum WordE235 WordE235;

enum WordE236 {
	WordE236RelativeHorizontalPositionMargin = '\002|\000\000',
	WordE236RelativeHorizontalPositionPage = '\002|\000\001',
	WordE236RelativeHorizontalPositionColumn = '\002|\000\002',
	WordE236RelativeHorizontalPositionCharacter = '\002|\000\003',
	WordE236RelativeHorizontalPositionLeftMargin = '\002|\000\004',
	WordE236RelativeHorizontalPositionRightMargin = '\002|\000\005',
	WordE236RelativeHorizontalPositionInnerMargin = '\002|\000\006',
	WordE236RelativeHorizontalPositionOuterMargin = '\002|\000\007'
};
typedef enum WordE236 WordE236;

enum WordE237 {
	WordE237RelativeVerticalPositionMargin = '\002}\000\000',
	WordE237RelativeVerticalPositionPage = '\002}\000\001',
	WordE237RelativeVerticalPositionParagraph = '\002}\000\002',
	WordE237RelativeVerticalPositionLine = '\002}\000\003',
	WordE237RelativeVerticalPositionTopMargin = '\002}\000\004',
	WordE237RelativeVerticalPositionBottomMargin = '\002}\000\005',
	WordE237RelativeVerticalPositionInnerMargin = '\002}\000\006',
	WordE237RelativeVerticalPositionOuterMargin = '\002}\000\007'
};
typedef enum WordE237 WordE237;

enum WordE238 {
	WordE238Help = '\002~\000\000',
	WordE238HelpAbout = '\002~\000\001',
	WordE238HelpContents = '\002~\000\003',
	WordE238HelpIndex = '\002~\000\005',
	WordE238HelpPsshelp = '\002~\000\007',
	WordE238HelpSearch = '\002~\000\011'
};
typedef enum WordE238 WordE238;

enum WordE239 {
	WordE239KeyCategoryNil = '\002~\377\377',
	WordE239KeyCategoryDisable = '\002\177\000\000',
	WordE239KeyCategoryCommand = '\002\177\000\001',
	WordE239KeyCategoryMacro = '\002\177\000\002',
	WordE239KeyCategoryFont = '\002\177\000\003',
	WordE239KeyCategoryAutoText = '\002\177\000\004',
	WordE239KeyCategoryStyle = '\002\177\000\005',
	WordE239KeyCategorySymbol = '\002\177\000\006',
	WordE239KeyCategoryPrefix = '\002\177\000\007'
};
typedef enum WordE239 WordE239;

enum WordE240 {
	WordE240No_key = '\002\200\000\377',
	WordE240Shift_key = '\002\200\002\000',
	WordE240Control_key = '\002\200\020\000',
	WordE240Command_key = '\002\200\001\000',
	WordE240Option_key = '\002\200\010\000',
	WordE240KeyAlt = '\002\200\010\000',
	WordE240A_key = '\002\200\000A',
	WordE240B_key = '\002\200\000B',
	WordE240C_key = '\002\200\000C',
	WordE240D_key = '\002\200\000D',
	WordE240E_key = '\002\200\000E',
	WordE240F_key = '\002\200\000F',
	WordE240G_key = '\002\200\000G',
	WordE240H_key = '\002\200\000H',
	WordE240I_Key = '\002\200\000I',
	WordE240J_key = '\002\200\000J',
	WordE240K_key = '\002\200\000K',
	WordE240L_key = '\002\200\000L',
	WordE240M_key = '\002\200\000M',
	WordE240N_key = '\002\200\000N',
	WordE240O_key = '\002\200\000O',
	WordE240P_key = '\002\200\000P',
	WordE240Q_key = '\002\200\000Q',
	WordE240R_key = '\002\200\000R',
	WordE240S_key = '\002\200\000S',
	WordE240T_key = '\002\200\000T',
	WordE240U_kwy = '\002\200\000U',
	WordE240V_key = '\002\200\000V',
	WordE240W_key = '\002\200\000W',
	WordE240X_key = '\002\200\000X',
	WordE240Y_key = '\002\200\000Y',
	WordE240Z_key = '\002\200\000Z',
	WordE240Key_number_0 = '\002\200\0000',
	WordE240Key_number_1 = '\002\200\0001',
	WordE240Key_numbe_2 = '\002\200\0002',
	WordE240Key_numbe_3 = '\002\200\0003',
	WordE240Key_number_4 = '\002\200\0004',
	WordE240Key_number_5 = '\002\200\0005',
	WordE240Key_number_6 = '\002\200\0006',
	WordE240KeyNumber_7 = '\002\200\0007',
	WordE240Key_number_8 = '\002\200\0008',
	WordE240Key_number_9 = '\002\200\0009',
	WordE240Backspace_key = '\002\200\000\010',
	WordE240Tab_key = '\002\200\000\011',
	WordE240Key_numeric_5_special = '\002\200\000\014',
	WordE240Return_key = '\002\200\000\015',
	WordE240Pause_key = '\002\200\000\023',
	WordE240Esc_key = '\002\200\000\033',
	WordE240Spacebar_key = '\002\200\000 ',
	WordE240Page_up_key = '\002\200\000!',
	WordE240Page_down_key = '\002\200\000\"',
	WordE240End_key = '\002\200\000#',
	WordE240Home_key = '\002\200\000$',
	WordE240Insert_key = '\002\200\000-',
	WordE240Delete_key = '\002\200\000.',
	WordE240Key_numeric_0 = '\002\200\000`',
	WordE240Key_numeric_1 = '\002\200\000a',
	WordE240Key_numeric_2 = '\002\200\000b',
	WordE240Key_numeric_3 = '\002\200\000c',
	WordE240Key_numeric_4 = '\002\200\000d',
	WordE240Key_numeric_5 = '\002\200\000e',
	WordE240Key_numeric_6 = '\002\200\000f',
	WordE240Key_numeric_7 = '\002\200\000g',
	WordE240Key_numeric_8 = '\002\200\000h',
	WordE240Key_numeric_9 = '\002\200\000i',
	WordE240Key_numeric_multiply = '\002\200\000j',
	WordE240Key_numeric_add = '\002\200\000k',
	WordE240Key_numeric_subtract = '\002\200\000m',
	WordE240Key_numeric_decimal = '\002\200\000n',
	WordE240Key_numeric_divide = '\002\200\000o',
	WordE240F1_key = '\002\200\000p',
	WordE240F2_key = '\002\200\000q',
	WordE240F3_key = '\002\200\000r',
	WordE240F4_key = '\002\200\000s',
	WordE240F5_key = '\002\200\000t',
	WordE240F6_key = '\002\200\000u',
	WordE240F7_key = '\002\200\000v',
	WordE240F8_key = '\002\200\000w',
	WordE240F9_key = '\002\200\000x',
	WordE240F10_key = '\002\200\000y',
	WordE240F11_key = '\002\200\000z',
	WordE240F12_key = '\002\200\000{',
	WordE240F13_key = '\002\200\000|',
	WordE240F14_key = '\002\200\000}',
	WordE240F15_key = '\002\200\000~',
	WordE240F16_key = '\002\200\000\177',
	WordE240Scroll_lock_key = '\002\200\000\221',
	WordE240Semi_colon_key = '\002\200\000\272',
	WordE240Equals_key = '\002\200\000\273',
	WordE240Comma_key = '\002\200\000\274',
	WordE240Hyphen_key = '\002\200\000\275',
	WordE240Period_key = '\002\200\000\276',
	WordE240Slash_key = '\002\200\000\277',
	WordE240Back_single_quote_key = '\002\200\000\300',
	WordE240Open_square_brace_key = '\002\200\000\333',
	WordE240Back_slash_key = '\002\200\000\334',
	WordE240Close_square_brace_key = '\002\200\000\335',
	WordE240Single_quote_key = '\002\200\000\336'
};
typedef enum WordE240 WordE240;

enum WordE241 {
	WordE241Olelink = '\002\201\000\000',
	WordE241Oleembed = '\002\201\000\001',
	WordE241Olecontrol = '\002\201\000\002'
};
typedef enum WordE241 WordE241;

enum WordE242 {
	WordE242OleverbPrimary = '\002\202\000\000',
	WordE242OleverbShow = '\002\201\377\377',
	WordE242OleverbOpen = '\002\201\377\376',
	WordE242OleverbHide = '\002\201\377\375',
	WordE242OleverbUiactivate = '\002\201\377\374',
	WordE242OleverbInPlaceActivate = '\002\201\377\373',
	WordE242OleverbDiscardUndoState = '\002\201\377\372'
};
typedef enum WordE242 WordE242;

enum WordE243 {
	WordE243InLine = '\002\203\000\000',
	WordE243FloatOverText = '\002\203\000\001'
};
typedef enum WordE243 WordE243;

enum WordE244 {
	WordE244LeftPortrait = '\002\204\000\000',
	WordE244CenterPortrait = '\002\204\000\001',
	WordE244RightPortrait = '\002\204\000\002',
	WordE244LeftLandscape = '\002\204\000\003',
	WordE244CenterLandscape = '\002\204\000\004',
	WordE244RightLandscape = '\002\204\000\005',
	WordE244LeftClockwise = '\002\204\000\006',
	WordE244CenterClockwise = '\002\204\000\007',
	WordE244RightClockwise = '\002\204\000\010'
};
typedef enum WordE244 WordE244;

enum WordE245 {
	WordE245FullBlock = '\002\205\000\000',
	WordE245ModifiedBlock = '\002\205\000\001',
	WordE245SemiBlock = '\002\205\000\002'
};
typedef enum WordE245 WordE245;

enum WordE246 {
	WordE246LetterTop = '\002\206\000\000',
	WordE246LetterBottom = '\002\206\000\001',
	WordE246LetterLeft = '\002\206\000\002',
	WordE246LetterRight = '\002\206\000\003'
};
typedef enum WordE246 WordE246;

enum WordE247 {
	WordE247SalutationInformal = '\002\207\000\000',
	WordE247SalutationFormal = '\002\207\000\001',
	WordE247SalutationBusiness = '\002\207\000\002',
	WordE247SalutationOther = '\002\207\000\003'
};
typedef enum WordE247 WordE247;

enum WordE248 {
	WordE248GenderFemale = '\002\210\000\000',
	WordE248GenderMale = '\002\210\000\001',
	WordE248GenderNeutral = '\002\210\000\002',
	WordE248GenderUnknown = '\002\210\000\003'
};
typedef enum WordE248 WordE248;

enum WordE249 {
	WordE249ByMoving = '\002\211\000\000',
	WordE249BySelecting = '\002\211\000\001'
};
typedef enum WordE249 WordE249;

enum WordE250 {
	WordE250UndefinedConstant = '\000\230\226\177',
	WordE250ToggleConstant = '\000\230\226~',
	WordE250ForwardConstant = '\?\377\377\377',
	WordE250BackwardConstant = '\300\000\000\001',
	WordE250AutoPosition = '\000\000\000\000',
	WordE250FirstConstant = '\000\000\000\001',
	WordE250CreatorCode = 'MSWD'
};
typedef enum WordE250 WordE250;

enum WordE251 {
	WordE251PasteOleobject = '\002\213\000\000',
	WordE251PasteRtf = '\002\213\000\001',
	WordE251PasteText = '\002\213\000\002',
	WordE251PasteMetafilePicture = '\002\213\000\003',
	WordE251PasteBitmap = '\002\213\000\004',
	WordE251PasteDeviceIndependentBitmap = '\002\213\000\005',
	WordE251PasteHyperlink = '\002\213\000\007',
	WordE251PasteShape = '\002\213\000\010',
	WordE251PasteEnhancedMetafile = '\002\212\377\377',
	WordE251PasteStyledText = '\002\213\000\011',
	WordE251PasteHtml = '\002\213\000\012',
	WordE251PastePDF = '\002\213\000\013'
};
typedef enum WordE251 WordE251;

enum WordE252 {
	WordE252PrintDocumentContent = '\002\214\000\000',
	WordE252PrintProperties = '\002\214\000\001',
	WordE252PrintComments = '\002\214\000\002',
	WordE252PrintStyles = '\002\214\000\003',
	WordE252PrintAutoTextEntries = '\002\214\000\004',
	WordE252PrintKeyAssignments = '\002\214\000\005',
	WordE252PrintEnvelope = '\002\214\000\006'
};
typedef enum WordE252 WordE252;

enum WordE253 {
	WordE253PrintAllPages = '\002\215\000\000',
	WordE253PrintOddPagesOnly = '\002\215\000\001',
	WordE253PrintEvenPagesOnly = '\002\215\000\002'
};
typedef enum WordE253 WordE253;

enum WordE254 {
	WordE254PrintAllDocument = '\002\216\000\000',
	WordE254PrintSelection = '\002\216\000\001',
	WordE254PrintCurrentPage = '\002\216\000\002',
	WordE254PrintFromTo = '\002\216\000\003',
	WordE254PrintRangeOfPages = '\002\216\000\004'
};
typedef enum WordE254 WordE254;

enum WordE255 {
	WordE255Spelling = '\002\217\000\000',
	WordE255Grammar = '\002\217\000\001',
	WordE255Thesaurus = '\002\217\000\002',
	WordE255Hyphenation = '\002\217\000\003',
	WordE255SpellingComplete = '\002\217\000\004',
	WordE255SpellingCustom = '\002\217\000\005',
	WordE255SpellingLegal = '\002\217\000\006',
	WordE255SpellingMedical = '\002\217\000\007',
	WordE255HangulHanjaConversion = '\002\217\000\010',
	WordE255HangulHanjaConversionCustom = '\002\217\000\011'
};
typedef enum WordE255 WordE255;

enum WordE256 {
	WordE256SpellingWordTypeSpellWord = '\002\220\000\000',
	WordE256SpellingWordTypeWildcard = '\002\220\000\001',
	WordE256SpellingWordTypeAnagram = '\002\220\000\002'
};
typedef enum WordE256 WordE256;

enum WordE257 {
	WordE257SpellingCorrect = '\002\221\000\000',
	WordE257SpellingNotInDictionary = '\002\221\000\001',
	WordE257SpellingCapitalization = '\002\221\000\002'
};
typedef enum WordE257 WordE257;

enum WordE258 {
	WordE258SpellingError = '\002\222\000\000',
	WordE258GrammaticalError = '\002\222\000\001'
};
typedef enum WordE258 WordE258;

enum WordE259 {
	WordE259InlineShapeEmbeddedOleobject = '\002\223\000\001',
	WordE259InlineShapeLinkedOleobject = '\002\223\000\002',
	WordE259InlineShapePicture = '\002\223\000\003',
	WordE259InlineShapeLinkedPicture = '\002\223\000\004',
	WordE259InlineShapeOlecontrolObject = '\002\223\000\005',
	WordE259InlineShapeHorizontalLine = '\002\223\000\006',
	WordE259InlineShapePictureHorizontalLine = '\002\223\000\007',
	WordE259InlineShapeLinkedPictureHorizontalLine = '\002\223\000\010',
	WordE259InlineShapePictureBullet = '\002\223\000\011',
	WordE259InlineShapeScriptAnchor = '\002\223\000\012',
	WordE259InlineShapeOWSAnchor = '\002\223\000\013',
	WordE259InlineShapeChart = '\002\223\000\014',
	WordE259InlineShapeDiagram = '\002\223\000\015',
	WordE259InlineShapeLockedCanvas = '\002\223\000\016',
	WordE259InlineShapeSmartartGraphic = '\002\223\000\017'
};
typedef enum WordE259 WordE259;

enum WordE260 {
	WordE260Tiled = '\002\224\000\000',
	WordE260Icons = '\002\224\000\001'
};
typedef enum WordE260 WordE260;

enum WordE261 {
	WordE261SelectionStartActive = '\002\225\000\001',
	WordE261SelectionAtEol = '\002\225\000\002',
	WordE261SelectionOvertype = '\002\225\000\004',
	WordE261SelectionActive = '\002\225\000\010',
	WordE261SelectionReplace = '\002\225\000\020',
	WordE261SelectionInactive = '\002\225\000\000',
	WordE261SelectionStartActiveAndAtEol = '\002\225\000\003',
	WordE261SelectionStartActiveAndOvertype = '\002\225\000\005',
	WordE261SelectionAtEolAndOvertype = '\002\225\000\006',
	WordE261SelectionStartActiveAndAtEolAndOvertype = '\002\225\000\007',
	WordE261SelectionStartActiveAndActive = '\002\225\000\011',
	WordE261SelectionAtEolAndActive = '\002\225\000\012',
	WordE261SelectionStartActiveAndAtEolAndActive = '\002\225\000\013',
	WordE261SelectionOvertypeAndActive = '\002\225\000\014',
	WordE261SelectionStartActiveAndOvertypeAndActive = '\002\225\000\015',
	WordE261SelectionAtEolAndOvertypeAndActive = '\002\225\000\016',
	WordE261SelectionStartActiveAndAtEolAndOvertypeAndActive = '\002\225\000\017',
	WordE261SelectionStartActiveAndReplace = '\002\225\000\021',
	WordE261SelectionAtEolAndReplace = '\002\225\000\022',
	WordE261SelectionStartActiveAndAtEolAndReplace = '\002\225\000\023',
	WordE261SelectionOvertypeAndReplace = '\002\225\000\024',
	WordE261SelectionStartActiveAndOvertypeAndReplace = '\002\225\000\025',
	WordE261SelectionAtEolAndOvertypeAndReplace = '\002\225\000\026',
	WordE261SelectionStartActiveAndAtEolAndOvertypeAndReplace = '\002\225\000\027',
	WordE261SelectionActiveAndReplace = '\002\225\000\030',
	WordE261SelectionStartActiveAndActiveAndReplace = '\002\225\000\031',
	WordE261SelectionAtEolAndActiveAndReplace = '\002\225\000\032',
	WordE261SelectionStartActiveAndAtEolAndActiveAndReplace = '\002\225\000\033',
	WordE261SelectionOvertypeAndActiveAndReplace = '\002\225\000\034',
	WordE261SelectionStartActiveAndOvertypeAndActiveAndReplace = '\002\225\000\035',
	WordE261SelectionAtEolAndOvertypeAndActiveAndReplace = '\002\225\000\036',
	WordE261SelectionStartActiveAndAtEolAndOvertypeAndActiveAndReplace = '\002\225\000\037'
};
typedef enum WordE261 WordE261;

enum WordE262 {
	WordE262AutoVersionOff = '\002\226\000\000',
	WordE262AutoVersionOnClose = '\002\226\000\001'
};
typedef enum WordE262 WordE262;

enum WordE263 {
	WordE263OrganizerObjectStyles = '\002\227\000\000',
	WordE263OrganizerObjectAutoText = '\002\227\000\001',
	WordE263OrganizerObjectCommandBars = '\002\227\000\002'
};
typedef enum WordE263 WordE263;

enum WordE264 {
	WordE264MatchParagraphMark = '\002\231\000\017',
	WordE264MatchTabCharacter = '\002\230\000\011',
	WordE264MatchCommentMark = '\002\230\000\005',
	WordE264MatchAnyCharacter = '\002\231\000\?',
	WordE264MatchAnyDigit = '\002\231\000\037',
	WordE264MatchAnyLetter = '\002\231\000/',
	WordE264MatchCaretCharacter = '\002\230\000\013',
	WordE264MatchColumnBreak = '\002\230\000\016',
	WordE264MatchEmDash = '\002\230 \024',
	WordE264MatchEnDash = '\002\230 \023',
	WordE264MatchEndnoteMark = '\002\231\000\023',
	WordE264MatchField = '\002\230\000\023',
	WordE264MatchFootnoteMark = '\002\231\000\022',
	WordE264MatchGraphic = '\002\230\000\001',
	WordE264MatchManualLineBreak = '\002\231\000\017',
	WordE264MatchManualPageBreak = '\002\231\000\034',
	WordE264MatchNonbreakingHyphen = '\002\230\000\036',
	WordE264MatchNonbreakingSpace = '\002\230\000\240',
	WordE264MatchOptionalHyphen = '\002\230\000\037',
	WordE264MatchSectionBreak = '\002\231\000,',
	WordE264MatchWhiteSpace = '\002\231\000w'
};
typedef enum WordE264 WordE264;

enum WordE265 {
	WordE265FindStop = '\002\231\000\000',
	WordE265FindContinue = '\002\231\000\001',
	WordE265FindAsk = '\002\231\000\002'
};
typedef enum WordE265 WordE265;

enum WordE266 {
	WordE266ActiveEndAdjustedPageNumber = '\002\232\000\001',
	WordE266ActiveEndSectionNumber = '\002\232\000\002',
	WordE266ActiveEndPageNumber = '\002\232\000\003',
	WordE266NumberOfPagesInDocument = '\002\232\000\004',
	WordE266HorizontalPositionRelativeToPage = '\002\232\000\005',
	WordE266VerticalPositionRelativeToPage = '\002\232\000\006',
	WordE266HorizontalPositionRelativeToTextBoundary = '\002\232\000\007',
	WordE266VerticalPositionRelativeToTextBoundary = '\002\232\000\010',
	WordE266FirstCharacterColumnNumber = '\002\232\000\011',
	WordE266FirstCharacterLineNumber = '\002\232\000\012',
	WordE266FrameIsSelected = '\002\232\000\013',
	WordE266WithInTable = '\002\232\000\014',
	WordE266StartOfRangeRowNumber = '\002\232\000\015',
	WordE266End_ofRangeRowNumber = '\002\232\000\016',
	WordE266MaximumNumberOfRows = '\002\232\000\017',
	WordE266StartOfRangeColumnNumber = '\002\232\000\020',
	WordE266End_ofRangeColumnNumber = '\002\232\000\021',
	WordE266MaximumNumberOfColumns = '\002\232\000\022',
	WordE266ZoomPercentage = '\002\232\000\023',
	WordE266SelectionMode = '\002\232\000\024',
	WordE266InfoCapsLock = '\002\232\000\025',
	WordE266InfoNumLock = '\002\232\000\026',
	WordE266OverType = '\002\232\000\027',
	WordE266RevisionMarking = '\002\232\000\030',
	WordE266InFootnoteEndnotePane = '\002\232\000\031',
	WordE266InCommentPane = '\002\232\000\032',
	WordE266InHeaderFooter = '\002\232\000\034',
	WordE266AtEndOfRowMarker = '\002\232\000\037',
	WordE266ReferenceOfType = '\002\232\000 ',
	WordE266HeaderFooterType = '\002\232\000!',
	WordE266InMasterDocument = '\002\232\000\"',
	WordE266InFootnote = '\002\232\000#',
	WordE266InEndnote = '\002\232\000$',
	WordE266InWordMail = '\002\232\000%',
	WordE266InClipboard = '\002\232\000&'
};
typedef enum WordE266 WordE266;

enum WordE267 {
	WordE267WrapSquare = '\002\233\000\000',
	WordE267WrapTight = '\002\233\000\001',
	WordE267WrapThrough = '\002\233\000\002',
	WordE267WrapNone = '\002\233\000\003',
	WordE267WrapTopBottom = '\002\233\000\004'
};
typedef enum WordE267 WordE267;

enum WordE312 {
	WordE312PictureWrapTypeInlineWithText = '\002\310\000\000',
	WordE312PictureWrapTypeSquare = '\002\310\000\001',
	WordE312PictureWrapTypeTight = '\002\310\000\002',
	WordE312PictureWrapTypeBehindText = '\002\310\000\003',
	WordE312PictureWrapTypeInFrontOfText = '\002\310\000\004',
	WordE312PictureWrapTypeThrough = '\002\310\000\005',
	WordE312PictureWrapTypeTopAndBottom = '\002\310\000\006'
};
typedef enum WordE312 WordE312;

enum WordE268 {
	WordE268WrapBoth = '\002\234\000\000',
	WordE268WrapLeft = '\002\234\000\001',
	WordE268WrapRight = '\002\234\000\002',
	WordE268WrapLargest = '\002\234\000\003'
};
typedef enum WordE268 WordE268;

enum WordE269 {
	WordE269OutlineLevel1 = '\002\235\000\001',
	WordE269OutlineLevel2 = '\002\235\000\002',
	WordE269OutlineLevel3 = '\002\235\000\003',
	WordE269OutlineLevel4 = '\002\235\000\004',
	WordE269OutlineLevel5 = '\002\235\000\005',
	WordE269OutlineLevel6 = '\002\235\000\006',
	WordE269OutlineLevel7 = '\002\235\000\007',
	WordE269OutlineLevel8 = '\002\235\000\010',
	WordE269OutlineLevel9 = '\002\235\000\011',
	WordE269OutlineLevelBodyText = '\002\235\000\012'
};
typedef enum WordE269 WordE269;

enum WordE270 {
	WordE270TextOrientationHorizontal = '\002\236\000\000',
	WordE270TextOrientationUpward = '\002\236\000\002',
	WordE270TextOrientationDownward = '\002\236\000\003',
	WordE270TextOrientationVerticalEastAsian = '\002\236\000\001',
	WordE270TextOrientationHorizontalRotatedEastAsian = '\002\236\000\004'
};
typedef enum WordE270 WordE270;

enum WordE271 {
	WordE271ArtNone = '\002\237\000\000',
	WordE271ArtApples = '\002\237\000\001',
	WordE271ArtMapleMuffins = '\002\237\000\002',
	WordE271ArtCakeSlice = '\002\237\000\003',
	WordE271ArtCandyCorn = '\002\237\000\004',
	WordE271ArtIceCreamCones = '\002\237\000\005',
	WordE271ArtChampagneBottle = '\002\237\000\006',
	WordE271ArtPartyGlass = '\002\237\000\007',
	WordE271ArtChristmasTree = '\002\237\000\010',
	WordE271ArtTrees = '\002\237\000\011',
	WordE271ArtPalmsColor = '\002\237\000\012',
	WordE271ArtBalloons3Colors = '\002\237\000\013',
	WordE271ArtBalloonsHotAir = '\002\237\000\014',
	WordE271ArtPartyFavor = '\002\237\000\015',
	WordE271ArtConfettiStreamers = '\002\237\000\016',
	WordE271ArtHearts = '\002\237\000\017',
	WordE271ArtHeartBalloon = '\002\237\000\020',
	WordE271ArtStars3D = '\002\237\000\021',
	WordE271ArtStarsShadowed = '\002\237\000\022',
	WordE271ArtStars = '\002\237\000\023',
	WordE271ArtSun = '\002\237\000\024',
	WordE271ArtEarth2 = '\002\237\000\025',
	WordE271ArtEarth1 = '\002\237\000\026',
	WordE271ArtPeopleHats = '\002\237\000\027',
	WordE271ArtSombrero = '\002\237\000\030',
	WordE271ArtPencils = '\002\237\000\031',
	WordE271ArtPackages = '\002\237\000\032',
	WordE271ArtClocks = '\002\237\000\033',
	WordE271ArtFirecrackers = '\002\237\000\034',
	WordE271ArtRings = '\002\237\000\035',
	WordE271ArtMapPins = '\002\237\000\036',
	WordE271ArtConfetti = '\002\237\000\037',
	WordE271ArtCreaturesButterfly = '\002\237\000 ',
	WordE271ArtCreaturesLadyBug = '\002\237\000!',
	WordE271ArtCreaturesFish = '\002\237\000\"',
	WordE271ArtBirdsFlight = '\002\237\000#',
	WordE271ArtScaredCat = '\002\237\000$',
	WordE271ArtBats = '\002\237\000%',
	WordE271ArtFlowersRoses = '\002\237\000&',
	WordE271ArtFlowersRedRose = '\002\237\000\'',
	WordE271ArtPoinsettias = '\002\237\000(',
	WordE271ArtHolly = '\002\237\000)',
	WordE271ArtFlowersTiny = '\002\237\000*',
	WordE271ArtFlowersPansy = '\002\237\000+',
	WordE271ArtFlowersModern2 = '\002\237\000,',
	WordE271ArtFlowersModern1 = '\002\237\000-',
	WordE271ArtWhiteFlowers = '\002\237\000.',
	WordE271ArtVine = '\002\237\000/',
	WordE271ArtFlowersDaisies = '\002\237\0000',
	WordE271ArtFlowersBlockPrint = '\002\237\0001',
	WordE271ArtDecoArchColor = '\002\237\0002',
	WordE271ArtFans = '\002\237\0003',
	WordE271ArtFilm = '\002\237\0004',
	WordE271ArtLightning1 = '\002\237\0005',
	WordE271ArtCompass = '\002\237\0006',
	WordE271ArtDoubleD = '\002\237\0007',
	WordE271ArtClassicalWave = '\002\237\0008',
	WordE271ArtShadowedSquares = '\002\237\0009',
	WordE271ArtTwistedLines1 = '\002\237\000:',
	WordE271ArtWaveline = '\002\237\000;',
	WordE271ArtQuadrants = '\002\237\000<',
	WordE271ArtCheckedBarColor = '\002\237\000=',
	WordE271ArtSwirligig = '\002\237\000>',
	WordE271ArtPushPinNote1 = '\002\237\000\?',
	WordE271ArtPushPinNote2 = '\002\237\000@',
	WordE271ArtPumpkin1 = '\002\237\000A',
	WordE271ArtEggsBlack = '\002\237\000B',
	WordE271ArtCup = '\002\237\000C',
	WordE271ArtHeartGray = '\002\237\000D',
	WordE271ArtGingerbreadMan = '\002\237\000E',
	WordE271ArtBabyPacifier = '\002\237\000F',
	WordE271ArtBabyRattle = '\002\237\000G',
	WordE271ArtCabins = '\002\237\000H',
	WordE271ArtHouseFunky = '\002\237\000I',
	WordE271ArtStarsBlack = '\002\237\000J',
	WordE271ArtSnowflakes = '\002\237\000K',
	WordE271ArtSnowflakeFancy = '\002\237\000L',
	WordE271ArtSkyrocket = '\002\237\000M',
	WordE271ArtSeattle = '\002\237\000N',
	WordE271ArtMusicNotes = '\002\237\000O',
	WordE271ArtPalmsBlack = '\002\237\000P',
	WordE271ArtMapleLeaf = '\002\237\000Q',
	WordE271ArtPaperClips = '\002\237\000R',
	WordE271ArtShorebirdTracks = '\002\237\000S',
	WordE271ArtPeople = '\002\237\000T',
	WordE271ArtPeopleWaving = '\002\237\000U',
	WordE271ArtEclipsingSquares2 = '\002\237\000V',
	WordE271ArtHypnotic = '\002\237\000W',
	WordE271ArtDiamondsGray = '\002\237\000X',
	WordE271ArtDecoArch = '\002\237\000Y',
	WordE271ArtDecoBlocks = '\002\237\000Z',
	WordE271ArtCirclesLines = '\002\237\000[',
	WordE271ArtPapyrus = '\002\237\000\\',
	WordE271ArtWoodwork = '\002\237\000]',
	WordE271ArtWeavingBraid = '\002\237\000^',
	WordE271ArtWeavingRibbon = '\002\237\000_',
	WordE271ArtWeavingAngles = '\002\237\000`',
	WordE271ArtArchedScallops = '\002\237\000a',
	WordE271ArtSafari = '\002\237\000b',
	WordE271ArtCelticKnotwork = '\002\237\000c',
	WordE271ArtCrazyMaze = '\002\237\000d',
	WordE271ArtEclipsingSquares1 = '\002\237\000e',
	WordE271ArtBirds = '\002\237\000f',
	WordE271ArtFlowersTeacup = '\002\237\000g',
	WordE271ArtNorthwest = '\002\237\000h',
	WordE271ArtSouthwest = '\002\237\000i',
	WordE271ArtTribal6 = '\002\237\000j',
	WordE271ArtTribal4 = '\002\237\000k',
	WordE271ArtTribal3 = '\002\237\000l',
	WordE271ArtTribal2 = '\002\237\000m',
	WordE271ArtTribal5 = '\002\237\000n',
	WordE271ArtXillusions = '\002\237\000o',
	WordE271ArtZanyTriangles = '\002\237\000p',
	WordE271ArtPyramids = '\002\237\000q',
	WordE271ArtPyramidsAbove = '\002\237\000r',
	WordE271ArtConfettiGrays = '\002\237\000s',
	WordE271ArtConfettiOutline = '\002\237\000t',
	WordE271ArtConfettiWhite = '\002\237\000u',
	WordE271ArtMosaic = '\002\237\000v',
	WordE271ArtLightning2 = '\002\237\000w',
	WordE271ArtHeebieJeebies = '\002\237\000x',
	WordE271ArtLightBulb = '\002\237\000y',
	WordE271ArtGradient = '\002\237\000z',
	WordE271ArtTriangleParty = '\002\237\000{',
	WordE271ArtTwistedLines2 = '\002\237\000|',
	WordE271ArtMoons = '\002\237\000}',
	WordE271ArtOvals = '\002\237\000~',
	WordE271ArtDoubleDiamonds = '\002\237\000\177',
	WordE271ArtChainLink = '\002\237\000\200',
	WordE271ArtTriangles = '\002\237\000\201',
	WordE271ArtTribal1 = '\002\237\000\202',
	WordE271ArtMarqueeToothed = '\002\237\000\203',
	WordE271ArtSharksTeeth = '\002\237\000\204',
	WordE271ArtSawtooth = '\002\237\000\205',
	WordE271ArtSawtoothGray = '\002\237\000\206',
	WordE271ArtPostageStamp = '\002\237\000\207',
	WordE271ArtWeavingStrips = '\002\237\000\210',
	WordE271ArtZigZag = '\002\237\000\211',
	WordE271ArtCrossStitch = '\002\237\000\212',
	WordE271ArtGems = '\002\237\000\213',
	WordE271ArtCirclesRectangles = '\002\237\000\214',
	WordE271ArtCornerTriangles = '\002\237\000\215',
	WordE271ArtCreaturesInsects = '\002\237\000\216',
	WordE271ArtZigZagStitch = '\002\237\000\217',
	WordE271ArtCheckered = '\002\237\000\220',
	WordE271ArtCheckedBarBlack = '\002\237\000\221',
	WordE271ArtMarquee = '\002\237\000\222',
	WordE271ArtBasicWhiteDots = '\002\237\000\223',
	WordE271ArtBasicWideMidline = '\002\237\000\224',
	WordE271ArtBasicWideOutline = '\002\237\000\225',
	WordE271ArtBasicWideInline = '\002\237\000\226',
	WordE271ArtBasicThinLines = '\002\237\000\227',
	WordE271ArtBasicWhiteDashes = '\002\237\000\230',
	WordE271ArtBasicWhiteSquares = '\002\237\000\231',
	WordE271ArtBasicBlackSquares = '\002\237\000\232',
	WordE271ArtBasicBlackDashes = '\002\237\000\233',
	WordE271ArtBasicBlackDots = '\002\237\000\234',
	WordE271ArtStarsTop = '\002\237\000\235',
	WordE271ArtCertificateBanner = '\002\237\000\236',
	WordE271ArtHandmade1 = '\002\237\000\237',
	WordE271ArtHandmade2 = '\002\237\000\240',
	WordE271ArtTornPaper = '\002\237\000\241',
	WordE271ArtTornPaperBlack = '\002\237\000\242',
	WordE271ArtCouponCutoutDashes = '\002\237\000\243',
	WordE271ArtCouponCutoutDots = '\002\237\000\244'
};
typedef enum WordE271 WordE271;

enum WordE272 {
	WordE272BorderDistanceFromText = '\002\240\000\000',
	WordE272BorderDistanceFromPageEdge = '\002\240\000\001'
};
typedef enum WordE272 WordE272;

enum WordE273 {
	WordE273ReplaceNone = '\002\241\000\000',
	WordE273ReplaceOne = '\002\241\000\001',
	WordE273ReplaceAll = '\002\241\000\002'
};
typedef enum WordE273 WordE273;

enum WordE274 {
	WordE274FontBiasDoNotCare = '\002\242\000\377',
	WordE274FontBiasDefault = '\002\242\000\000',
	WordE274FontBiasEastAsian = '\002\242\000\001'
};
typedef enum WordE274 WordE274;

enum WordE275 {
	WordE275BrowserLevelV4 = '\002\243\000\000',
	WordE275BrowserLevelMicrosoftInternetExplorer5 = '\002\243\000\001'
};
typedef enum WordE275 WordE275;

enum WordE276 {
	WordE276EnclosureCircle = '\002\244\000\000',
	WordE276EnclosureSquare = '\002\244\000\001',
	WordE276EnclosureTriangle = '\002\244\000\002',
	WordE276EnclosureDiamond = '\002\244\000\003'
};
typedef enum WordE276 WordE276;

enum WordE277 {
	WordE277EncloseStyleNone = '\002\245\000\000',
	WordE277EncloseStyleSmall = '\002\245\000\001',
	WordE277EncloseStyleLarge = '\002\245\000\002'
};
typedef enum WordE277 WordE277;

enum WordE278 {
	WordE278LayoutModeDefault = '\002\246\000\000',
	WordE278LayoutModeGrid = '\002\246\000\001',
	WordE278LayoutModeLineGrid = '\002\246\000\002',
	WordE278LayoutModeGenko = '\002\246\000\003'
};
typedef enum WordE278 WordE278;

enum WordE279 {
	WordE279ForAEmailMessage = '\002\247\000\000',
	WordE279ForADocument = '\002\247\000\001',
	WordE279ForAWebPage = '\002\247\000\002'
};
typedef enum WordE279 WordE279;

enum WordE308 {
	WordE308TwoLinesInOneNone = '\002\304\000\000',
	WordE308TwoLinesInOneNoBrackets = '\002\304\000\001',
	WordE308TwoLinesInOneParentheses = '\002\304\000\002',
	WordE308TwoLinesInOneSquareBrackets = '\002\304\000\003',
	WordE308TwoLinesInOneAngleBrackets = '\002\304\000\004',
	WordE308TwoLinesInOneCurlyBrackets = '\002\304\000\005'
};
typedef enum WordE308 WordE308;

enum WordE309 {
	WordE309HorizontalInVerticalNone = '\002\305\000\000',
	WordE309HorizontalInVerticalFitInLine = '\002\305\000\001',
	WordE309HorizontalInVerticalResizeLine = '\002\305\000\002'
};
typedef enum WordE309 WordE309;

enum WordE280 {
	WordE280HorizontalLineAlignLeft = '\002\250\000\000',
	WordE280HorizontalLineAlignCenter = '\002\250\000\001',
	WordE280HorizontalLineAlignRight = '\002\250\000\002'
};
typedef enum WordE280 WordE280;

enum WordE281 {
	WordE281HorizontalLinePercentWidth = '\002\250\377\377',
	WordE281HorizontalLineFixedWidth = '\002\250\377\376'
};
typedef enum WordE281 WordE281;

enum WordE310 {
	WordE310PhoneticGuideAlignmentCenter = '\002\306\000\000',
	WordE310PhoneticGuideAlignmentZeroOneZero = '\002\306\000\001',
	WordE310PhoneticGuideAlignmentOneTwoOne = '\002\306\000\002',
	WordE310PhoneticGuideAlignmentLeft = '\002\306\000\003',
	WordE310PhoneticGuideAlignmentRight = '\002\306\000\004',
	WordE310PhoneticGuideAlignmentRightVertical = '\002\306\000\005'
};
typedef enum WordE310 WordE310;

enum WordE282 {
	WordE282TableDirectionRtl = '\002\252\000\000',
	WordE282TableDirectionLtr = '\002\252\000\001'
};
typedef enum WordE282 WordE282;

enum WordE283 {
	WordE283GutterPositionLeft = '\002\253\000\000',
	WordE283GutterPositionTop = '\002\253\000\001',
	WordE283GutterPositionRight = '\002\253\000\002'
};
typedef enum WordE283 WordE283;

enum WordE284 {
	WordE284GutterStyleLatin = '\002\253\377\366',
	WordE284GutterStyleBidi = '\002\254\000\002',
	WordE284GutterStyleNone = '\002\254\000\000'
};
typedef enum WordE284 WordE284;

enum WordE285 {
	WordE285ShapeTop = '\002\235\275\301',
	WordE285ShapeLeft = '\002\235\275\302',
	WordE285ShapeBottom = '\002\235\275\303',
	WordE285ShapeRight = '\002\235\275\304',
	WordE285ShapeCenter = '\002\235\275\305',
	WordE285ShapeInside = '\002\235\275\306',
	WordE285ShapeOutside = '\002\235\275\307'
};
typedef enum WordE285 WordE285;

enum WordE286 {
	WordE286TableTop = '\002\236\275\301',
	WordE286TableLeft = '\002\236\275\302',
	WordE286TableBottom = '\002\236\275\303',
	WordE286TableRight = '\002\236\275\304',
	WordE286TableCenter = '\002\236\275\305',
	WordE286TableInside = '\002\236\275\306',
	WordE286TableOutside = '\002\236\275\307'
};
typedef enum WordE286 WordE286;

enum WordE287 {
	WordE287Word8TableBehavior = '\002\257\000\000',
	WordE287Word9TableBehavior = '\002\257\000\001'
};
typedef enum WordE287 WordE287;

enum WordE288 {
	WordE288AutoFitFixed = '\002\260\000\000',
	WordE288AutoFitContent = '\002\260\000\001',
	WordE288AutoFitWindow = '\002\260\000\002'
};
typedef enum WordE288 WordE288;

enum WordE289 {
	WordE289Word8ListBehavior = '\002\261\000\000',
	WordE289Word9ListBehavior = '\002\261\000\001',
	WordE289Word10ListBehavior = '\002\261\000\002'
};
typedef enum WordE289 WordE289;

enum WordE290 {
	WordE290PreferredWidthAuto = '\002\262\000\001',
	WordE290PreferredWidthPercent = '\002\262\000\002',
	WordE290PreferredWidthPoints = '\002\262\000\003'
};
typedef enum WordE290 WordE290;

enum WordE291 {
	WordE291NewBlankDocument = '\002\263\000\000',
	WordE291NewWebPage = '\002\263\000\001',
	WordE291NewNotebookDocument = '\002\263\000\004',
	WordE291NewPublishingDocument = '\002\263\000\005'
};
typedef enum WordE291 WordE291;

enum WordE292 {
	WordE292UserFirst = '\002\264\000\001',
	WordE292UserLast = '\002\264\000\002',
	WordE292UserCompany = '\002\264\000\024',
	WordE292UserWorkStreet = '\002\264\000\026',
	WordE292UserWorkCity = '\002\264\000\027',
	WordE292UserWorkState = '\002\264\000\030',
	WordE292UserWorkZip = '\002\264\000\031',
	WordE292UserWorkPhone = '\002\264\000\035',
	WordE292UserEmailAddress1 = '\002\264\000f'
};
typedef enum WordE292 WordE292;

enum WordE293 {
	WordE293PasteDefault = '\002\265\000\000',
	WordE293SingleCellText = '\002\265\000\005',
	WordE293SingleCellTable = '\002\265\000\006',
	WordE293ListContinueNumbering = '\002\265\000\007',
	WordE293ListRestartNumbering = '\002\265\000\010',
	WordE293TableInsertAsRows = '\002\265\000\013',
	WordE293TableAppendTable = '\002\265\000\012',
	WordE293TableOriginalFormatting = '\002\265\000\014',
	WordE293ChartPicture = '\002\265\000\015',
	WordE293Chart = '\002\265\000\016',
	WordE293ChartLinked = '\002\265\000\017',
	WordE293FormatOriginalFormatting = '\002\265\000\020',
	WordE293FormatSurroundingFormattingWithEmphasis = '\002\265\000\024',
	WordE293FormatPlainText = '\002\265\000\026',
	WordE293TableOverwriteCells = '\002\265\000\027',
	WordE293ListCombineWithExistingList = '\002\265\000\030',
	WordE293ListDontMerge = '\002\265\000\031'
};
typedef enum WordE293 WordE293;

enum WordE294 {
	WordE294GoForward = '\002\266\000\001',
	WordE294GoBackward = '\002\266\000\002',
	WordE294ANumericConstant = '\002\266\000\000'
};
typedef enum WordE294 WordE294;

enum WordE311 {
	WordE311LineEndingCrLf = '\002\307\000\000',
	WordE311LineEndingCrOnly = '\002\307\000\001',
	WordE311LineEndingLfOnly = '\002\307\000\002',
	WordE311LineEndingLfCr = '\002\307\000\003',
	WordE311LineEndingLsPs = '\002\307\000\004'
};
typedef enum WordE311 WordE311;

enum WordE299 {
	WordE299ConditionFirstRow = '\002\301\000\000',
	WordE299ConditionLastRow = '\002\301\000\001',
	WordE299ConditionOddRowBanding = '\002\301\000\002',
	WordE299ConditionEvenRowBanding = '\002\301\000\003',
	WordE299ConditionFirstColumn = '\002\301\000\004',
	WordE299ConditionLastColumn = '\002\301\000\005',
	WordE299ConditionOddColumnBanding = '\002\301\000\006',
	WordE299ConditionEvenColumnBanding = '\002\301\000\007',
	WordE299ConditionTopRightCell = '\002\301\000\010',
	WordE299ConditionTopLeftCell = '\002\301\000\011',
	WordE299ConditionBottomRightCell = '\002\301\000\012',
	WordE299ConditionBottomLeftCell = '\002\301\000\013'
};
typedef enum WordE299 WordE299;

enum WordE295 {
	WordE295UnitALine = '\002\267\000\005',
	WordE295UnitAStory = '\002\267\000\006',
	WordE295UnitAScreen = '\002\267\000\007',
	WordE295UnitASection = '\002\267\000\010',
	WordE295UnitAColumn = '\002\267\000\011',
	WordE295UnitARow = '\002\267\000\012'
};
typedef enum WordE295 WordE295;

enum WordE296 {
	WordE296HighlightOn = '\002\270\377\377',
	WordE296HighlightOff = '\002\270\000\000',
	WordE296ANumericConstant = '\002\270\000\000'
};
typedef enum WordE296 WordE296;

enum WordE297 {
	WordE297CompareTargetSelected = '\002\271\000\000',
	WordE297CompareTargetCurrent = '\002\271\000\001',
	WordE297CompareTargetNew = '\002\271\000\002'
};
typedef enum WordE297 WordE297;

enum WordE298 {
	WordE298MergeTargetSelected = '\002\272\000\000',
	WordE298MergeTargetCurrent = '\002\272\000\001',
	WordE298MergeTargetNew = '\002\272\000\002'
};
typedef enum WordE298 WordE298;

enum WordE300 {
	WordE300RevisionsViewFinal = '\002\274\000\000',
	WordE300RevisionsViewOriginal = '\002\274\000\001'
};
typedef enum WordE300 WordE300;

enum WordE301 {
	WordE301RevisionsViewBalloons = '\002\275\000\000',
	WordE301RevisionsViewInline = '\002\275\000\001'
};
typedef enum WordE301 WordE301;

enum WordE303 {
	WordE303BalloonPrintOrientationAutomatic = '\002\277\000\000',
	WordE303BalloonPrintOrientationPreserve = '\002\277\000\001',
	WordE303BalloonPrintOrientationLandscape = '\002\277\000\002'
};
typedef enum WordE303 WordE303;

enum WordE304 {
	WordE304BalloonMarginLeft = '\002\300\000\000',
	WordE304BalloonMarginRight = '\002\300\000\001'
};
typedef enum WordE304 WordE304;

enum WordE331 {
	WordE331MinorVersion = '\002\333\000\000',
	WordE331MajorVersion = '\002\333\000\001',
	WordE331OverwriteCurrentVersion = '\002\333\000\002'
};
typedef enum WordE331 WordE331;

enum WordE332 {
	WordE332LockNone = '\002\334\000\000',
	WordE332LockReservation = '\002\334\000\001',
	WordE332LockEphemeral = '\002\334\000\002',
	WordE332LockChanged = '\002\334\000\003'
};
typedef enum WordE332 WordE332;

enum Word4017 {
	Word4017FieldOptions = 'w298',
	Word4017Field = 'w170'
};
typedef enum Word4017 Word4017;

enum Word4018 {
	Word4018FieldOptions = 'w298',
	Word4018Field = 'w170'
};
typedef enum Word4018 Word4018;

enum Word4024 {
	Word4024Revision = 'w219',
	Word4024Conflict = 'o120'
};
typedef enum Word4024 Word4024;

enum Word4025 {
	Word4025Revision = 'w219',
	Word4025Conflict = 'o120'
};
typedef enum Word4025 Word4025;

enum Word4013 {
	Word4013FootnoteOptions = 'WopN',
	Word4013EndnoteOptions = 'WopE'
};
typedef enum Word4013 Word4013;

enum Word4014 {
	Word4014FootnoteOptions = 'WopN',
	Word4014EndnoteOptions = 'WopE'
};
typedef enum Word4014 Word4014;

enum Word4015 {
	Word4015FootnoteOptions = 'WopN',
	Word4015EndnoteOptions = 'WopE'
};
typedef enum Word4015 Word4015;

enum Word4004 {
	Word4004Document = 'docu',
	Word4004Window = 'cwin',
	Word4004Pane = 'w120'
};
typedef enum Word4004 Word4004;

enum Word4019 {
	Word4019Font = 'w137',
	Word4019Frame = 'w175',
	Word4019SelectionObject = 'WSoj'
};
typedef enum Word4019 Word4019;

enum Word4021 {
	Word4021Field = 'w170',
	Word4021Frame = 'w175',
	Word4021FormField = 'w177',
	Word4021DataMergeField = 'w187',
	Word4021SelectionObject = 'WSoj',
	Word4021PageNumber = 'w225'
};
typedef enum Word4021 Word4021;

enum Word4023 {
	Word4023ListFormat = 'w123',
	Word4023WordList = 'w236'
};
typedef enum Word4023 Word4023;

enum Word4002 {
	Word4002Application = 'capp',
	Word4002Document = 'docu'
};
typedef enum Word4002 Word4002;

enum Word4003 {
	Word4003Application = 'capp',
	Word4003Document = 'docu'
};
typedef enum Word4003 Word4003;

enum Word4011 {
	Word4011Find = 'w124',
	Word4011Replacement = 'w125',
	Word4011SelectionObject = 'WSoj'
};
typedef enum Word4011 Word4011;

enum Word4012 {
	Word4012DropCap = 'w133',
	Word4012TabStop = 'w135',
	Word4012TextInput = 'w178',
	Word4012KeyBinding = 'w242'
};
typedef enum Word4012 Word4012;

enum Word4010 {
	Word4010Document = 'docu',
	Word4010ListFormat = 'w123',
	Word4010WordList = 'w236'
};
typedef enum Word4010 Word4010;

enum Word4020 {
	Word4020Field = 'w170',
	Word4020Frame = 'w175',
	Word4020FormField = 'w177',
	Word4020DataMergeField = 'w187',
	Word4020SelectionObject = 'WSoj',
	Word4020PageNumber = 'w225'
};
typedef enum Word4020 Word4020;

enum Word4005 {
	Word4005Window = 'cwin',
	Word4005Pane = 'w120'
};
typedef enum Word4005 Word4005;

enum Word4009 {
	Word4009Document = 'docu',
	Word4009ListFormat = 'w123',
	Word4009WordList = 'w236'
};
typedef enum Word4009 Word4009;

enum Word4007 {
	Word4007Window = 'cwin',
	Word4007Pane = 'w120'
};
typedef enum Word4007 Word4007;

enum Word4001 {
	Word4001Application = 'capp',
	Word4001Document = 'docu',
	Word4001Window = 'cwin'
};
typedef enum Word4001 Word4001;

enum Word4008 {
	Word4008Document = 'docu',
	Word4008ListFormat = 'w123',
	Word4008WordList = 'w236'
};
typedef enum Word4008 Word4008;

enum Word4006 {
	Word4006Window = 'cwin',
	Word4006Pane = 'w120'
};
typedef enum Word4006 Word4006;

enum Word4022 {
	Word4022TableOfFigures = 'w184',
	Word4022TableOfContents = 'w198'
};
typedef enum Word4022 Word4022;

enum Word4016 {
	Word4016LinkFormat = 'w167',
	Word4016FieldOptions = 'w298',
	Word4016TableOfFigures = 'w184',
	Word4016TableOfContents = 'w198',
	Word4016TableOfAuthorities = 'w200',
	Word4016Dialog = 'w202',
	Word4016Index = 'w215'
};
typedef enum Word4016 Word4016;

enum WordE315 {
	WordE315E31502caffff = '\002\312\377\377',
	WordE315E31502cb0000 = '\002\313\000\000',
	WordE315E31502cb0001 = '\002\313\000\001',
	WordE315E31502cb0002 = '\002\313\000\002',
	WordE315E31502cb0003 = '\002\313\000\003',
	WordE315E31502cb0004 = '\002\313\000\004',
	WordE315E31502cb0005 = '\002\313\000\005',
	WordE315E31502cb0006 = '\002\313\000\006',
	WordE315E31502cb0007 = '\002\313\000\007',
	WordE315E31502cb0008 = '\002\313\000\010',
	WordE315E31502cb0009 = '\002\313\000\011',
	WordE315E31502cb000a = '\002\313\000\012',
	WordE315E31502cb000b = '\002\313\000\013',
	WordE315E31502cb000c = '\002\313\000\014',
	WordE315E31502cb000d = '\002\313\000\015',
	WordE315E31502cb000e = '\002\313\000\016',
	WordE315E31502cb000f = '\002\313\000\017'
};
typedef enum WordE315 WordE315;

enum Word4027 {
	Word4027Shape = 'pShp',
	Word4027CalloutFormat = 'w275'
};
typedef enum Word4027 Word4027;

enum Word4028 {
	Word4028Callout = 'cD00',
	Word4028CalloutFormat = 'w275'
};
typedef enum Word4028 Word4028;

enum Word4029 {
	Word4029Callout = 'cD00',
	Word4029CalloutFormat = 'w275'
};
typedef enum Word4029 Word4029;

enum Word4030 {
	Word4030Callout = 'cD00',
	Word4030CalloutFormat = 'w275'
};
typedef enum Word4030 Word4030;

enum Word4031 {
	Word4031Shape = 'pShp',
	Word4031FillFormat = 'w278'
};
typedef enum Word4031 Word4031;

enum Word4032 {
	Word4032Shape = 'pShp',
	Word4032FillFormat = 'w278'
};
typedef enum Word4032 Word4032;

enum Word4033 {
	Word4033Shape = 'pShp',
	Word4033FillFormat = 'w278'
};
typedef enum Word4033 Word4033;

enum Word4034 {
	Word4034Shape = 'pShp',
	Word4034FillFormat = 'w278'
};
typedef enum Word4034 Word4034;

enum Word4035 {
	Word4035Shape = 'pShp',
	Word4035FillFormat = 'w278'
};
typedef enum Word4035 Word4035;

enum Word4036 {
	Word4036Shape = 'pShp',
	Word4036FillFormat = 'w278'
};
typedef enum Word4036 Word4036;

enum Word4037 {
	Word4037Shape = 'pShp',
	Word4037FillFormat = 'w278'
};
typedef enum Word4037 Word4037;

enum Word4038 {
	Word4038Shape = 'pShp',
	Word4038FillFormat = 'w278'
};
typedef enum Word4038 Word4038;

enum Word4039 {
	Word4039Shape = 'pShp',
	Word4039ThreeDFormat = 'w286'
};
typedef enum Word4039 Word4039;

enum Word4026 {
	Word4026Shape = 'pShp',
	Word4026InlineShape = 'w257'
};
typedef enum Word4026 Word4026;

enum Word4046 {
	Word4046Paragraph = 'cpar',
	Word4046ParagraphFormat = 'w136'
};
typedef enum Word4046 Word4046;

enum Word4040 {
	Word4040TextRange = 'w122',
	Word4040Section = 'w130',
	Word4040Paragraph = 'cpar',
	Word4040ParagraphFormat = 'w136',
	Word4040WordStyle = 'w173'
};
typedef enum Word4040 Word4040;

enum Word4041 {
	Word4041Paragraph = 'cpar',
	Word4041ParagraphFormat = 'w136'
};
typedef enum Word4041 Word4041;

enum Word4050 {
	Word4050Paragraph = 'cpar',
	Word4050ParagraphFormat = 'w136'
};
typedef enum Word4050 Word4050;

enum Word4051 {
	Word4051Paragraph = 'cpar',
	Word4051ParagraphFormat = 'w136'
};
typedef enum Word4051 Word4051;

enum Word4043 {
	Word4043Paragraph = 'cpar',
	Word4043ParagraphFormat = 'w136'
};
typedef enum Word4043 Word4043;

enum Word4042 {
	Word4042Paragraph = 'cpar',
	Word4042ParagraphFormat = 'w136'
};
typedef enum Word4042 Word4042;

enum Word4048 {
	Word4048Paragraph = 'cpar',
	Word4048ParagraphFormat = 'w136'
};
typedef enum Word4048 Word4048;

enum Word4047 {
	Word4047Paragraph = 'cpar',
	Word4047ParagraphFormat = 'w136'
};
typedef enum Word4047 Word4047;

enum Word4049 {
	Word4049Paragraph = 'cpar',
	Word4049ParagraphFormat = 'w136'
};
typedef enum Word4049 Word4049;

enum Word4044 {
	Word4044Paragraph = 'cpar',
	Word4044ParagraphFormat = 'w136'
};
typedef enum Word4044 Word4044;

enum Word4045 {
	Word4045Paragraph = 'cpar',
	Word4045ParagraphFormat = 'w136'
};
typedef enum Word4045 Word4045;

enum Word4058 {
	Word4058Row = 'crow',
	Word4058RowOptions = 'WrOp'
};
typedef enum Word4058 Word4058;

enum Word4059 {
	Word4059Row = 'crow',
	Word4059RowOptions = 'WrOp'
};
typedef enum Word4059 Word4059;

enum Word4060 {
	Word4060Column = 'ccol',
	Word4060ColumnOptions = 'WcOp'
};
typedef enum Word4060 Word4060;

enum Word4053 {
	Word4053Table = 'ctbl',
	Word4053Column = 'ccol'
};
typedef enum Word4053 Word4053;

enum Word4052 {
	Word4052Table = 'ctbl',
	Word4052Row = 'crow',
	Word4052Column = 'ccol',
	Word4052Cell = 'ccel',
	Word4052RowOptions = 'WrOp',
	Word4052ColumnOptions = 'WcOp',
	Word4052TableStyle = 'w292',
	Word4052Condition = 'w293'
};
typedef enum Word4052 Word4052;

enum Word4057 {
	Word4057Row = 'crow',
	Word4057Cell = 'ccel',
	Word4057RowOptions = 'WrOp'
};
typedef enum Word4057 Word4057;

enum Word4056 {
	Word4056Column = 'ccol',
	Word4056Cell = 'ccel',
	Word4056ColumnOptions = 'WcOp'
};
typedef enum Word4056 Word4056;

enum Word4054 {
	Word4054Table = 'ctbl',
	Word4054Column = 'ccol'
};
typedef enum Word4054 Word4054;

enum Word4055 {
	Word4055Table = 'ctbl',
	Word4055Column = 'ccol'
};
typedef enum Word4055 Word4055;



/*
 * Standard Suite
 */

// A scriptable object
@interface WordBaseObject : SBObject

@property (copy) NSDictionary *properties;  // All of the object's properties

- (void) closeSaving:(WordSavo)saving savingIn:(NSURL *)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) openFileName:(NSString *)fileName confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate Revert:(BOOL)Revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate fileConverter:(WordE162)fileConverter;  // Open the specified object(s). Returns a reference to the opened file when using the file name parameter.
- (void) printWithProperties:(WordPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(NSURL *)in_ as:(NSNumber *)as;  // Save an object
- (void) select;  // Make a selection
- (WordMathObject *) addMathArgumentBeforeArg:(NSInteger)beforeArg toFunction:(WordMathFunction *)toFunction toMatrixRow:(WordMathMatrixRow *)toMatrixRow toMatrixColumn:(WordMathMatrixColumn *)toMatrixColumn;  // Inserts an argument into an equation with variable number of arguments, i.e. math delimiter and math equation array objects, and returns a math object object. You must specify only one function, row, or argument.
- (void) autoPullLocksLocalDocument:(WordDocument *)localDocument serverDocument:(WordDocument *)serverDocument baseDocument:(WordDocument *)baseDocument;
- (void) autosaveDocument:(WordDocument *)document;  // Causes an autosave to occur for the given document or all documents if no document is specified.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Word can check out a specified document from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified document from a server to a local computer for editing. Returns a String that represents the local path and file name of the document checked out.
- (void) duplicatePage;  // Duplicates the page of the current selection and moves the selection to the new page. This command only works when in Publishing Layout View.
- (void) insertPage;  // Insert a page following the page of the current selection.
- (void) removePage;  // Removes the page of the selection and moves the selection to the following page. When removing the last page, the selection is moved to the previous page. This command only works when in Publishing Layout View.
- (void) threeWayMergeLocalDocument:(WordDocument *)localDocument serverDocument:(WordDocument *)serverDocument baseDocument:(WordDocument *)baseDocument favorSource:(BOOL)favorSource;
- (void) toggleShowCodes;

@end

// A basic application program
@interface WordBaseApplication : WordBaseObject

@property (readonly) BOOL frontmost;  // Is this the frontmost application?
@property (copy, readonly) NSString *name;  // the name
@property (copy, readonly) NSString *version;  // the version of the application


@end

// A basic document
@interface WordBaseDocument : WordBaseObject

@property (readonly) BOOL modified;  // Has the document been modified since the last save?
@property (copy, readonly) NSString *name;  // the name


@end

// A basic window
@interface WordBasicWindow : WordBaseObject

@property NSRect bounds;  // the boundary rectangle for the window
@property (readonly) BOOL closeable;  // Does the window have a close box?
@property (readonly) BOOL titled;  // Does the window have a title bar?
@property (readonly) NSInteger entryIndex;  // the number of the window
@property (readonly) BOOL floating;  // Does the window float?
@property (readonly) BOOL modal;  // Is the window modal?
@property NSPoint position;  // upper left coordinates of the window
@property (readonly) BOOL resizable;  // Is the window resizable?
@property (readonly) BOOL zoomable;  // Is the window zoomable?
@property BOOL zoomed;  // Is the window zoomed?
@property (copy, readonly) NSString *name;  // the title of the window
@property (readonly) BOOL visible;  // Is the window visible?
@property (readonly) BOOL collapsable;  // Is the window collapasable?
@property BOOL collapsed;  // Is the window collapsed?
@property (readonly) BOOL sheet;  // Is this window a sheet window?


@end

@interface WordPrintSettings : SBObject

@property NSInteger copies;  // the number of copies of a document to be printed 
@property BOOL collating;  // Should printed copies be collated?
@property NSInteger startingPage;  // the first page of the document to be printed
@property NSInteger endingPage;  // the last page of the document to be printed
@property NSInteger pagesAcross;  // number of logical pages laid across a physical page
@property NSInteger pagesDown;  // number of logical pages laid out down a physical page
@property WordEnum errorHandling;  // how errors are handled
@property (copy) NSString *faxNumber;  // for fax number
@property (copy) NSString *targetPrinter;  // the queue name of the target printer

- (void) closeSaving:(WordSavo)saving savingIn:(NSURL *)savingIn;  // Close an object
- (NSInteger) dataSizeAs:(NSNumber *)as;  // Return the size in bytes of an object
- (void) delete;  // Delete an element from an object
- (SBObject *) duplicateTo:(SBObject *)to;  // Duplicate object(s)
- (BOOL) exists;  // Verify if an object exists
- (SBObject *) moveTo:(SBObject *)to;  // Move object(s) to a new location
- (void) openFileName:(NSString *)fileName confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate Revert:(BOOL)Revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate fileConverter:(WordE162)fileConverter;  // Open the specified object(s). Returns a reference to the opened file when using the file name parameter.
- (void) printWithProperties:(WordPrintSettings *)withProperties printDialog:(BOOL)printDialog;  // Print the specified object(s)
- (void) saveIn:(NSURL *)in_ as:(NSNumber *)as;  // Save an object
- (void) select;  // Make a selection
- (WordMathObject *) addMathArgumentBeforeArg:(NSInteger)beforeArg toFunction:(WordMathFunction *)toFunction toMatrixRow:(WordMathMatrixRow *)toMatrixRow toMatrixColumn:(WordMathMatrixColumn *)toMatrixColumn;  // Inserts an argument into an equation with variable number of arguments, i.e. math delimiter and math equation array objects, and returns a math object object. You must specify only one function, row, or argument.
- (void) autoPullLocksLocalDocument:(WordDocument *)localDocument serverDocument:(WordDocument *)serverDocument baseDocument:(WordDocument *)baseDocument;
- (void) autosaveDocument:(WordDocument *)document;  // Causes an autosave to occur for the given document or all documents if no document is specified.
- (BOOL) canCheckOutFileName:(NSString *)fileName;  // Returns True if Word can check out a specified document from a server.
- (void) checkOutFileName:(NSString *)fileName;  // Copies a specified document from a server to a local computer for editing. Returns a String that represents the local path and file name of the document checked out.
- (void) duplicatePage;  // Duplicates the page of the current selection and moves the selection to the new page. This command only works when in Publishing Layout View.
- (void) insertPage;  // Insert a page following the page of the current selection.
- (void) removePage;  // Removes the page of the selection and moves the selection to the following page. When removing the last page, the selection is moved to the previous page. This command only works when in Publishing Layout View.
- (void) threeWayMergeLocalDocument:(WordDocument *)localDocument serverDocument:(WordDocument *)serverDocument baseDocument:(WordDocument *)baseDocument favorSource:(BOOL)favorSource;
- (void) toggleShowCodes;

@end



/*
 * Microsoft Office Suite
 */

// A control within a command bar.
@interface WordCommandBarControl : WordBaseObject

@property BOOL beginGroup;  // Returns or sets if the command bar control appears at the beginning of a group of controls on the command bar.
@property (readonly) BOOL builtIn;  // Returns true if the command bar control is a built-in command bar control.
@property (readonly) WordMCLT controlType;  // Returns the type of the command bar control.
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
@property WordMBns buttonState;  // Returns or set the appearance of a command bar button control.  The property is read-only for built-in command bar buttons.
@property WordMBTs buttonStyle;  // Returns or sets the way a command button control is displayed.
@property NSInteger faceId;  // Returns or sets the Id number for the face of the command bar button control.


@end

// A combobox menu control within a command bar.
@interface WordCommandBarCombobox : WordCommandBarControl

@property WordMXcb comboboxStyle;  // Returns or sets the way a command bar combobox control is displayed.
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

- (SBElementArray *) commandBarControls;


@end

// Toolbars used in all of the Office applications.
@interface WordCommandBar : WordBaseObject

- (SBElementArray *) commandBarControls;

@property WordMBPS barPosition;  // Returns or sets the position of the command bar.
@property (readonly) WordMBTY barType;  // Returns the type of this command bar.
@property (readonly) BOOL builtIn;  // True if the command bar is built-in.
@property (copy, readonly) NSString *context;  // Returns or sets a string that determines where a command bar will be saved.
@property (readonly) BOOL embeddable;  // Returns if the command bar can be displayed inside the document window.
@property BOOL embedded;  // Returns or sets if the command bar will be displayed inside the document window.
@property BOOL enabled;  // Returns or set if the command bar is enabled.
@property (readonly) NSInteger entry_index;  // The index of the command bar.
@property NSInteger height;  // Returns or sets the height of the command bar.
@property NSInteger leftPosition;  // Returns or sets the left position of the command bar.
@property (copy) NSString *localName;  // Returns or sets the name of the command bar in the localized language of the application.  An error is returned when trying to set the local name of a built-in command bar.
@property (copy) NSString *name;  // Returns or sets the name of the command bar.
@property (copy) NSArray *protection;  // Returns or sets the way a command bar is protected from user customization.  It accepts a list of the following items: no protection, no customize, no resize, no move, no change visible, no change dock, no vertical dock, no horizontal dock.
@property NSInteger rowIndex;  // Returns or sets the docking order of a command bar in relation to other command bars in the same docking area.  Can be an integer greater than zero.
@property NSInteger top;  // Returns or sets the top position of a command bar.
@property BOOL visible;  // Returns or sets if the command bar is visible.
@property NSInteger width;  // Returns or sets the width in pixels of the command bar.


@end

@interface WordDocumentProperty : WordBaseObject

@property (copy) NSNumber *documentPropertyType;  // Returns or sets the document property type.
@property (copy) NSString *linkSource;  // Returns or sets the source of a lined custom document property.
@property BOOL linkToContent;  // True if the value of the document property is lined to the content of the container document.  False if the value is static.  This only applies to custom document properties.  For built-in properties this is always false.
@property (copy) NSString *name;  // Returns or sets the name of the document property.
@property (copy) NSString *value;  // Returns or sets the value of the document property.


@end

@interface WordCustomDocumentProperty : WordDocumentProperty


@end

@interface WordWebPageFont : WordBaseObject

@property (copy) NSString *fixedWidthFont;  // Returns or sets the fixed-width font setting.
@property double fixedWidthFontSize;  // Returns or sets the fixed-width font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.
@property (copy) NSString *proportionalFont;  // Returns or sets the proportional font setting.
@property double proportionalFontSize;  // Returns or sets the proportional font size.  You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


@end



/*
 * Microsoft Word Suite
 */

// Represents a single comment.
@interface WordWordComment : WordBaseObject

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
@interface WordWordList : WordBaseObject

- (SBElementArray *) paragraphs;

@property (readonly) BOOL singleListTemplate;  // Returns if the entire Word list object uses the same list template.
@property (copy, readonly) NSString *styleName;
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the Word list.

- (void) applyListTemplateListTemplate:(WordListTemplate *)listTemplate continuePreviousList:(BOOL)continuePreviousList defaultListBehavior:(WordE289)defaultListBehavior;  // Applies a set of list-formatting characteristics to the specified Word list object.

@end

// Represents application and document options in Word.
@interface WordWordOptions : WordBaseObject

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
@property BOOL autoFormatAsYouTypeReplaceSymbols;  // Returns or sets if two consecutive hyphens -- are replaced with an en dash – or an em dash — as you type.
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
@property BOOL autoFormatReplaceSymbols;  // Returns or set if two consecutive hyphens -- are replaced by an en dash – or an em dash — when Word formats a document or range automatically.
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
@property WordE110 commentsColor;  // Returns or sets the color of comments.
@property BOOL confirmConversions;  // Returns or sets if Word displays the convert file dialog box before it opens or inserts a file that isn't a Word document or template. In the convert file dialog box, the user chooses the format to convert the file from.
@property BOOL convertHighAnsiToEastAsian;  // Returns or sets if Microsoft Word converts text that is associated with an East Asian font to the appropriate font when it opens a document.
@property BOOL createBackup;  // Returns or sets if Word creates a backup copy each time a document is saved.
@property BOOL dashMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search.
@property (copy) NSColor *defaultBorderColor;  // Returns or sets the default RGB color to use for new border objects.
@property WordE110 defaultBorderColorIndex;  // Returns or sets the default line color index for borders.
@property WordMCoI defaultBorderColorThemeIndex;  // Returns or sets the default line color for borders.
@property WordE167 defaultBorderLineStyle;  // Returns or sets the default border line style
@property WordE168 defaultBorderLineWidth;  // Returns or sets the default line width of borders.
@property WordE110 defaultHighlightColorIndex;  // Returns or sets the color index  used to highlight text formatted with the highlight button.
@property WordE162 defaultOpenFormat;  // Returns or sets the default file converter used to open documents.
@property WordE318 deletedCellColor;  // Returns or sets the color of cells that are deleted while change tracking is enabled.
@property WordE110 deletedTextColor;  // Returns or sets the color of text that is deleted while change tracking is enabled.
@property WordE227 deletedTextMark;  // Returns or sets the format of text that is deleted while change tracking is enabled.
@property BOOL displayGridLines;  // Returns or sets if Microsoft Word displays the document grid. 
@property BOOL displayPasteOptions;  // Returns or sets if Microsoft Word  displays the Paste Options button, which displays directly under newly pasted text.
@property BOOL dzMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL enableSound;  // Returns or sets if Word makes the computer respond with a sound whenever an error occurs.
@property (readonly) BOOL envelopeFeederInstalled;  // Returns true if the current printer has a special feeder for envelopes.
@property BOOL fancyFontMenu;  // Returns or sets if the fancy font menu is shown.
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
@property WordE318 insertedCellColor;  // Returns or sets the color of cells that are inserted while change tracking is enabled.
@property WordE110 insertedTextColor;  // Returns or sets the color of text that is inserted while change tracking is enabled.
@property WordE225 insertedTextMark;  // Returns or sets how Microsoft Word formats inserted text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the inserted text color property to control the look of inserted text.
@property BOOL iterationMarkMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between types of repetition marks during a search.
@property BOOL kanjiMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between standard and nonstandard kanji ideography during a search.
@property BOOL keepTrackOfFormatting;  // Returns or sets if Microsoft Word keeps track of formatting.
@property BOOL kiKuMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL liveWordCount;  // Returns or sets if the instant word count is displayed in the status bar.
@property BOOL mapPaperSize;  // Returns or sets if documents formatted for another country's or region's standard paper size, for example, A4 are automatically adjusted so that they're printed correctly on your country's/region's standard paper size, for example, Letter.
@property WordE171 measurementUnit;  // Returns or sets the standard measurement unit for Microsoft Word.
@property WordE318 mergedCellColor;  // Returns or sets the color of cells that are merged while change tracking is enabled.
@property WordE110 moveFromTextColor;  // Returns or sets the color of text that is moved from while change tracking is enabled.
@property WordE317 moveFromTextMark;  // Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text.
@property WordE110 moveToTextColor;  // Returns or sets the color of text that is moved to while change tracking is enabled.
@property WordE316 moveToTextMark;  // Returns or sets how Microsoft Word formats moved text while change tracking is enabled. If change tracking is not enabled, this property is ignored. Use this property with the moved text color property to control the look of moved text.
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
@property WordE312 pictureWrapTypes;  // Returns or sets the wrapping that is used to insert or paste pictures.
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
@property WordE110 revisedLinesColor;  // Returns or sets the color of changed lines in a document with tracked changes.
@property WordE226 revisedLinesMark;  // Returns or sets the placement of changed lines in a document with tracked changes.
@property WordE110 revisedPropertiesColor;  // Returns or sets the color index used to mark formatting changes while change tracking is enabled. 
@property WordE228 revisedPropertiesMark;  // Returns or sets the mark used to show formatting changes while change tracking is enabled.
@property NSInteger saveInterval;  // Returns or sets the time interval in minutes for saving autorecover information.
@property BOOL saveNormalPrompt;  // Returns or sets if Microsoft Word prompts the user for confirmation to save changes to the Normal template before it quits. if this is set to false Word automatically saves changes to the Normal template before it quits.
@property BOOL savePropertiesPrompt;  // Returns or sets if Microsoft Word prompts for document property information when saving a new document.
@property BOOL sendMailAttach;  // True if the send to command on the file menu inserts the active document as an attachment to a mail message. False if the send to command inserts the contents of the active document as text in a mail message.
@property BOOL showReadabilityStatistics;  // Returns or sets if Microsoft Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar.
@property BOOL showWizardWelcome;  // Returns or sets if the welcome wizard should be shown.
@property BOOL smallKanaMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between diphthongs and double consonants during a search.
@property BOOL smartCutPaste;  // Returns or sets if Microsoft Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs.
@property BOOL smartParagraphSelection;  // Returns or sets if Microsoft Word includes the paragraph mark in a selection when selecting most or all of a paragraph.
@property BOOL snapToGrid;  // Returns or sets if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized in new documents.
@property BOOL snapToShapes;  // Returns or sets if Word automatically aligns autoshapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other autoshapes or East Asian characters in new documents.
@property BOOL spaceMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between space markers used during a search.
@property WordE318 splitCellColor;  // Returns or sets the color of cells that are split while change tracking is enabled.
@property BOOL suggestFromMainDictionaryOnly;  // Returns or sets if Microsoft Word draws spelling suggestions from the main dictionary only. If false it draws spelling suggestions from the main dictionary and any custom dictionaries that have been added.
@property BOOL suggestSpellingCorrections;  // Returns or sets if Microsoft Word always suggests alternative spellings for each misspelled word when checking spelling.
@property BOOL tabIndentKey;  // Returns or sets if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered and centered paragraphs to left-aligned.
@property BOOL tcMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.
@property BOOL trackFormatting;  // Returns or sets tracking formatting flag.
@property BOOL trackMoves;  // Returns or sets tracking moves flag.
@property BOOL updateFieldsAtPrint;  // Returns or sets if Microsoft Word updates fields automatically before printing a document. 
@property BOOL updateLinksAtOpen;  // Returns or sets if Microsoft Word automatically updates all embedded OLE links in a document when it's opened.
@property BOOL updateLinksAtPrint;  // Returns or sets if Microsoft Word updates fields automatically before printing a document.
@property BOOL useCharacterUnit;  // Returns or sets if Microsoft Word uses characters as the default measurement unit for the current document.
@property BOOL useGermanSpellingReform;  // Returns or sets if Microsoft Word uses the German post-reform spelling rules when checking spelling.
@property BOOL warnBeforeSavingPrintingSendingMarkup;  // Returns or sets if Microsoft Word displays a warning when saving, printing, or sending as e-mail a document containing comments or tracked changes.
@property BOOL zjMatchFuzzy;  // Returns or sets if Microsoft Word ignores the distinction between some Japanese characters.


@end

// Represents a single add-in, either installed or not installed.
@interface WordAddIn : WordBaseObject

@property (readonly) BOOL autoload;  // Returns true if the specified add-in is automatically loaded when Word is started.
@property (readonly) BOOL compiled;  // Returns true if the specified add-in is a Word add-in library. False if the add-in is a template.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property BOOL installed;  // Returns or sets if the specified add-in is installed.
@property (copy, readonly) NSString *name;  // Returns the name of the add in.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified add-in.


@end

// An AppleScript object representing the Microsoft Word application.
@interface WordApplication : SBApplication

- (SBElementArray *) documents;
- (SBElementArray *) windows;
- (SBElementArray *) recentFiles;
- (SBElementArray *) fileConverters;
- (SBElementArray *) captionLabels;
- (SBElementArray *) addIns;
- (SBElementArray *) commandBars;
- (SBElementArray *) templates;
- (SBElementArray *) keyBindings;
- (SBElementArray *) dictionaries;

@property BOOL Word51Menus;  // Returns or sets if Word 5.1 menus should be used.
@property (copy, readonly) WordDocument *activeDocument;  // Returns the currently active document object.
@property (copy) NSString *activePrinter;  // Returns or sets the name of the active printer
@property (copy, readonly) WordWindow *activeWindow;  // Returns the currently active window object.
@property (copy, readonly) NSString *applicationVersion;  // Returns the Microsoft Word version number as a string.
@property (copy, readonly) WordAutocorrect *autocorrectObject;  // Returns the autocorrect object
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
@property WordE138 displayAlerts;  // Returns or sets the way certain alerts or messages are handled while an AppleScript is running.
@property BOOL displayAutoCompleteTips;  // Returns or set if Microsoft Word displays tips that suggest text for completing words, dates, or phrases as you type.
@property BOOL displayRecentFiles;  // Returns or sets if the names of recently used files are displayed on the file menu.
@property BOOL displayRibbon;  // Returns or sets if Word displays the Ribbon in at least one document window.  Setting this property to true will display the Ribbon in all windows.  Setting this property to false turns off all Ribbons in all windows.
@property BOOL displayScreenTips;  // Returns or set if comments, footnotes, endnotes, and hyperlinks are displayed as tips.  Text marked as having comments is highlighted.
@property BOOL displayScrollBars;  // Returns or sets if Word displays a scroll bar in at least one document window.  Setting this property to true will display horizontal and vertical scrollbars in all windows.  Setting this property to false turns off all scrollbars in all windows.
@property BOOL displayStatusBar;  // Returns or sets if the status bar is displayed.
@property BOOL doPrintPreview;  // Returns or set if print preview is the current view.
@property (copy, readonly) NSArray *fontNames;  // Returns the list of names of all of the available fonts
@property NSInteger keyboardScript;  // Returns or sets the current keyboard script
@property (copy, readonly) NSArray *landscapeFontNames;  // Returns the list of names of all of the available landscape fonts
@property (readonly) NSInteger localizedLanguage;  // Returns the Windows language ID for the locale that Microsoft Word is using.
@property (copy, readonly) WordMailingLabel *mailingLabelObject;  // Returns the mailing label object.
@property (copy, readonly) WordMathAutocorrect *mathAc;  // Returns a math autocorrect object that represents the auto correct entries for equations.
@property (copy, readonly) NSString *name;  // Returns the name of this application.
@property (copy, readonly) WordTemplate *normalTemplate;  // Returns the normal template object
@property (readonly) BOOL numLock;  // Returns the state of the num lock key.  True if the keys on the numeric keypad insert numbers, false if the keys move the insertion point.
@property (copy, readonly) NSString *path;  // Returns the path to the application
@property (copy, readonly) NSString *pathSeparator;  // Returns the character used to separate folder names.
@property (copy, readonly) NSArray *portraitFontNames;  // Returns the list of names of all of the available portrait fonts
@property BOOL ribbonExpanded;  // Returns or sets if the Ribbon in expanded.
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

- (void) quitSaving:(WordSavo)saving;  // Quit an application program
- (void) select;  // Make a selection
- (id) doScript:(NSString *)x;  // Execute a script
- (void) WordHelpHelpType:(WordE238)helpType;  // Displays on-line Help information.
- (void) accept:(Word4024)x;  // Accepts the specified tracked change. The revision marks are removed, and the change is incorporated into the document.
- (void) activateObject:(Word4004)x;  // Activate this object.
- (WordAddIn *) addAddinFileName:(NSString *)fileName install:(BOOL)install;  // Returns an add in object that represents an add-in added to the list of available add-ins.
- (WordCoauthLock *) addCoauthLockToRange:(WordTextRange *)toRange toParagraph:(WordParagraph *)toParagraph toColumn:(WordColumn *)toColumn toCell:(WordCell *)toCell toRow:(WordRow *)toRow toTable:(WordTable *)toTable toSelection:(WordSelectionObject *)toSelection lockType:(WordE332)lockType inRange:(WordSelectionObject *)inRange inCoauthoring:(WordCoauthoring *)inCoauthoring inCoauthor:(WordCoauthor *)inCoauthor;  // Add a co-authoring lock.
- (void) arrangeWindowsArrangeStyle:(WordE260)arrangeStyle;  // Arrange windows on the screen.
- (NSInteger) buildKeyCodeKey1:(WordE240)key1 key2:(WordE240)key2 key3:(WordE240)key3 key4:(WordE240)key4;  // Returns a unique number for the specified key combination.
- (WordE102) canContinuePreviousList:(Word4023)x listTemplate:(WordListTemplate *)listTemplate;  // Returns whether the formatting from the previous list can be continued.
- (double) centimetersToPointsCentimeters:(double)centimeters;  // Converts a measurement from centimeters to points.
- (void) changeFileOpenDirectoryPath:(NSString *)path;  // Sets the folder in which Microsoft Word searches for documents.
- (BOOL) checkGrammar:(Word4002)x textToCheck:(NSString *)textToCheck;  // Checks a string for grammatical errors. Returns a boolean to indicate whether the string contains grammatical errors. True if the string contains no errors.
- (BOOL) checkSpelling:(Word4003)x textToCheck:(NSString *)textToCheck customDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Checks a string for spelling errors. Returns true if the string has no spelling errors.
- (NSString *) cleanStringItemToCheck:(NSString *)itemToCheck;  // Removes nonprinting characters character codes 1 – 29 and special Microsoft Word characters from the specified string or changes them to spaces character code 32. Returns the result as a string.
- (void) clear:(Word4012)x;  // Removes the object.
- (void) clearFormatting:(Word4011)x;  // Removes text and paragraph formatting from a selection or from the formatting specified in a find or replace operation.
- (void) convertNumbersToText:(Word4009)x numberType:(WordE158)numberType;  // Changes the list numbers and listnum fields in the document object to text.
- (void) copyObject:(Word4020)x;  // Copies the content of the object to the clipboard.
- (NSInteger) countNumberedItems:(Word4010)x numberType:(WordE158)numberType level:(NSInteger)level;  // Returns the number of bulleted or numbered items and listnum fields in the document object.
- (WordDocument *) createNewDocumentAttachedTemplate:(WordTemplate *)attachedTemplate newTemplate:(BOOL)newTemplate newDocumentType:(WordE291)newDocumentType;  // Create a new document
- (void) createNewEquationFromRange:(WordTextRange *)fromRange inDocument:(WordDocument *)inDocument inRange:(WordSelectionObject *)inRange inSelection:(WordSelectionObject *)inSelection;  // Creates an equation, from the text equation contained within the specified range, and returns a text range object that contains the new equation.
- (void) createNewFieldTextRange:(WordTextRange *)textRange fieldType:(WordE183)fieldType fieldText:(NSString *)fieldText preserveFormatting:(BOOL)preserveFormatting;  // Create a new field
- (void) cutObject:(Word4021)x;  // Removes the object from the document and places it on the clipboard.
- (BOOL) doWordRepeatTimes:(NSInteger)times;  // Repeats the most recent editing action one or more times. Returns true if the commands were repeated successfully.
- (WordKeyBinding *) findKeyKeyCode:(NSInteger)keyCode key_code_2:(WordE240)key_code_2;  // Returns a key binding object that represents the specified key combination
- (NSString *) getDefaultFilePathFilePathType:(WordE230)filePathType;  // Returns the default folders for items such as documents, templates, and graphics.
- (WordDialog *) getDialog:(WordE186)x;  // Returns a dialog object for the specified dialog.
- (NSString *) getInternationalInformation:(WordE115)x;  // Get the specified international information
- (NSArray *) getKeysBoundToKeyCategory:(WordE239)keyCategory command:(NSString *)command;  // Returns a list key binding objects that represents all the key combinations assigned to the specified item.
- (WordListGallery *) getListGallery:(WordE150)x;  // Returns the specified list gallery object object.
- (NSDictionary *) getSpellingSuggestionsItemToCheck:(NSString *)itemToCheck customDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary suggestionMode:(WordE256)suggestionMode customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Returns a record that specifies the error kind and a list of words.  The AEKeyword for the error kind is type and AEKeyword for the list is list.
- (WordSynonymInfo *) getSynonymInfoObjectItemToCheck:(NSString *)itemToCheck languageID:(WordE182)languageID;  // Returns a synonym info object that contains the information from the thesaurus on the synonyms, antonyms, or related words and expressions for the specified word or phrase.
- (NSString *) getUserPropertyPropertyType:(WordE292)propertyType;
- (WordWebPageFont *) getWebpageFont:(WordMChS)x;  // Return the specified web page font object for a given character set.
- (double) inchesToPointsInches:(double)inches;  // Converts a measurement from inches to points.
- (void) insertText:(NSString *)text at:(SBObject *)at;  // Insert the string at the specified location.
- (void) insertAutoTextAt:(WordTextRange *)at;  // Attempts to match the text in the specified range or the text surrounding the range with an existing auto text entry name.  If any such match is found, the auto text entry is inserted.  If no match an error occurs.
- (void) insertBreakAt:(WordTextRange *)at breakType:(WordE169)breakType;  // Inserts a break in the specified place of the specified kind.
- (void) insertCaptionAt:(WordTextRange *)at captionLabel:(WordE210)captionLabel title:(NSString *)title captionPosition:(WordE117)captionPosition;  // Inserts a caption immediately preceding or following the specified range.
- (void) insertCrossReferenceAt:(WordTextRange *)at referenceType:(WordE211)referenceType referenceKind:(WordE212)referenceKind referenceItem:(NSString *)referenceItem insertAsHyperlink:(BOOL)insertAsHyperlink includePosition:(BOOL)includePosition;  // Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined.
- (void) insertDatabaseAt:(WordTextRange *)at format:(WordE180)format style:(NSInteger)style linkToSource:(BOOL)linkToSource connection:(NSString *)connection SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1 passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate dataSource:(NSString *)dataSource from:(NSInteger)from to:(NSInteger)to includeFields:(BOOL)includeFields;  // Retrieves data from a data source and inserts the data as a table in place of the specified range.
- (void) insertDateTimeAt:(WordTextRange *)at dateTimeFormat:(NSString *)dateTimeFormat insertAsField:(BOOL)insertAsField;  // Insert the correct date or time, or both, either as text or as a time field at the specified location.
- (void) insertFileAt:(WordTextRange *)at fileName:(NSString *)fileName fileRange:(NSString *)fileRange confirmConversions:(BOOL)confirmConversions link:(BOOL)link;
- (void) insertParagraphAt:(WordTextRange *)at;  // Replaces the specified range with a new paragraph.  If you do not want to replace the range, use the collapse range method before using this method.
- (void) insertSymbolAt:(WordTextRange *)at characterNumber:(NSInteger)characterNumber font:(NSString *)font unicode:(BOOL)unicode bias:(WordE274)bias;  // Inserts a symbol in place of the specified range.  If you do not want to replace the range, use the collapse range method before using this method.
- (NSString *) keyStringKeyCode:(NSInteger)keyCode key_code_2:(WordE240)key_code_2;  // Returns the key combination string for the specified keys 
- (void) largeScroll:(Word4005)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls a window by the specified number of screens. This method is equivalent to clicking just before or just after the scroll boxes on the horizontal and vertical scroll bars.
- (double) linesToPointsLines:(double)lines;  // Converts a measurement from lines to points. 1 line = 12 points.
- (void) listCommandsListAllCommands:(BOOL)listAllCommands;  // Creates a new document and then inserts a table of Microsoft Word commands along with their associated shortcut keys and menu assignments.
- (double) millimetersToPointsMillimeters:(double)millimeters;  // Converts a measurement from millimeters to points.
- (void) organizerCopySource:(NSString *)source destination:(NSString *)destination name:(NSString *)name organizerObjectType:(WordE263)organizerObjectType;  // Copies the specified autotext entry, toolbar, style, or macro project item from the source document or template to the destination document or template.
- (void) organizerDeleteSource:(NSString *)source name:(NSString *)name organizerObjectType:(WordE263)organizerObjectType;  // Deletes the specified style, autotext entry, toolbar, or macro project item from a document or template.
- (void) organizerRenameSource:(NSString *)source name:(NSString *)name newName:(NSString *)newName organizerObjectType:(WordE263)organizerObjectType;  // Renames the specified style, autotext entry, toolbar, or macro project item in a document or template.
- (void) pageScroll:(Word4007)x down:(NSInteger)down up:(NSInteger)up;  // Scrolls through the window page by page.
- (double) picasToPointsPicas:(double)picas;  // Converts a measurement from picas to points.
- (double) pointsToCentimetersPoints:(double)points;  // Converts a measurement from points to centimeters.
- (double) pointsToInchesPoints:(double)points;  // Converts a measurement from points to inches.
- (double) pointsToLinesPoints:(double)points;  // Converts a measurement from points to lines. 1 line = 12 points.
- (double) pointsToMillimetersPoints:(double)points;  // Converts a measurement from points to millimeters.
- (double) pointsToPicasPoints:(double)points;  // Converts a measurement from points to picas.
- (void) printOut:(Word4001)x append:(BOOL)append printOutRange:(WordE254)printOutRange outputFileName:(NSString *)outputFileName pageFrom:(NSInteger)pageFrom pageTo:(NSInteger)pageTo printOutItem:(WordE252)printOutItem printCopies:(NSInteger)printCopies printOutPageType:(WordE253)printOutPageType printToFile:(BOOL)printToFile collate:(BOOL)collate fileName:(NSString *)fileName manualDuplexPrint:(BOOL)manualDuplexPrint;  // Prints out all or part of the specified or active document. This command has been deprecated; use the Print command in the Standard Suite.
- (void) reject:(Word4025)x;  // Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.
- (void) removeNumbers:(Word4008)x numberType:(WordE158)numberType;  // Removes numbers or bullets from the document
- (void) remveEmphemeralLocksInRange:(WordSelectionObject *)inRange inCoauthoring:(WordCoauthoring *)inCoauthoring inCoauthor:(WordCoauthor *)inCoauthor;
- (void) resetContinuationNotice:(Word4015)x;  // Resets the footnote or endnote continuation notice to the default notice. The default notice is blank.
- (void) resetContinuationSeparator:(Word4014)x;  // Resets the footnote or endnote continuation separator to the default separator. The default separator is a long horizontal line that separates document text from notes continued from the previous page.
- (void) resetIgnoreAll;  // Clears the list of words that were previously ignored during a spelling check. After you run this method, previously ignored words are checked along with all the other words.
- (void) resetSeparator:(Word4013)x;  // Resets the footnote or endnote separator to the default separator. The default separator is a short horizontal line that separates document text from notes.
- (WordLanguage *) retrieveLanguage:(WordE182)x;  // Returns the language object for the specified language.
- (void) runVBMacroMacroName:(NSString *)macroName;  // Runs a Visual Basic macro.
- (void) screenRefresh;  // Updates the display on the monitor with the current information in the video memory buffer. You can use this method after using the screen updating property to disable screen updates.
- (void) setDefaultFilePathFilePathType:(WordE230)filePathType path:(NSString *)path;  // Sets the default folders for items such as documents, templates, and graphics.
- (void) setUserPropertyPropertyType:(WordE292)propertyType propertyValue:(NSString *)propertyValue;
- (void) smallScroll:(Word4006)x down:(NSInteger)down up:(NSInteger)up toRight:(NSInteger)toRight toLeft:(NSInteger)toLeft;  // Scrolls a window by the specified number of lines. This method is equivalent to clicking the scroll arrows on the horizontal and vertical scroll bars.
- (void) substituteFontUnavailableFont:(NSString *)unavailableFont substituteFont:(NSString *)substituteFont;  // Sets font-mapping options, which are reflected in the font substitution dialog box
- (void) unlink:(WordField*)x;
- (void) unloadAddinsRemoveFromList:(BOOL)removeFromList;  // Unloads all loaded add-ins and, depending on the value of the remove from list argument, removes them from the list of add-ins.
- (void) update:(Word4016)x;  // Updates the values shown in a built-in Microsoft Word dialog box, updates the specified link, or updates the entries shown in specified index, table of authorities, table of figures or table of contents.
- (void) updatePageNumbers:(Word4022)x;  // Updates the page numbers for items in the specified table of contents or table of figures.
- (void) updateSource:(Word4018)x;
- (void) automaticLength:(Word4027)x;  // Specifies that the first segment of the callout line the segment attached to the text callout box be scaled automatically when the callout is moved. Applies only to callouts whose lines consist of more than one segment.
- (void) customDrop:(Word4028)x drop:(double)drop;  // Sets the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
- (void) customLength:(Word4029)x length:(double)length;  // Specifies that the first segment of the callout, line the segment attached to the text callout box, retain a fixed length whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment.
- (void) oneColorGradient:(Word4031)x gradientStyle:(WordMGdS)gradientStyle gradientVariant:(NSInteger)gradientVariant gradientDegree:(double)gradientDegree;  // Sets the specified fill to a one-color gradient.
- (void) patterned:(Word4032)x pattern:(WordPpTy)pattern;  // Sets the specified fill to a pattern.
- (void) presetDrop:(Word4030)x DropType:(WordMCDt)DropType;  // Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.
- (void) presetGradient:(Word4033)x style:(WordMGdS)style gradientVariant:(NSInteger)gradientVariant presetGradientType:(WordMPGb)presetGradientType;  // Sets the specified fill to a preset gradient.
- (void) presetTextured:(Word4034)x presetTexture:(WordMPzT)presetTexture;  // Sets the specified fill to a preset texture
- (void) resetRotation:(Word4039)x;  // Resets the extrusion rotation around the x-axis and the y-axis to zero so that the front of the extrusion faces forward. This method doesn't reset the rotation around the z-axis.
- (void) saveAsPicture:(Word4026)x pictureType:(WordMPiT)pictureType fileName:(NSString *)fileName;  // Saves the shape in the requested file using the stated graphic format
- (void) solid:(Word4035)x;  // Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.
- (void) twoColorGradient:(Word4036)x gradientStyle:(WordMGdS)gradientStyle gradientVariant:(NSInteger)gradientVariant;  // Sets the specified fill to a two-color gradient.
- (void) userPicture:(Word4037)x pictureFile:(NSString *)pictureFile;  // Fills the specified shape with one large image. If you want to fill the shape with small tiles of an image, use the user textured method.
- (void) userTextured:(Word4038)x textureFile:(NSString *)textureFile;  // Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the user picture method.
- (void) closeUp:(Word4041)x;  // Removes any spacing before the specified paragraphs.
- (WordBorder *) getBorder:(Word4040)x whichBorder:(WordE122)whichBorder;  // Returns the specified border object.
- (void) indentCharWidth:(Word4050)x count:(NSInteger)count;  // Indents one or more paragraphs by a specified number of characters.
- (void) indentFirstLineCharWidth:(Word4051)x count:(NSInteger)count;  // Indents the first line of one or more paragraphs by a specified number of characters.
- (void) openOrCloseUp:(Word4043)x;  // If spacing before the specified paragraphs is zero, this method sets spacing to 12 points. If spacing before the paragraphs is greater than zero, this method sets spacing to zero.
- (void) openUp:(Word4042)x;  // Sets spacing before the specified paragraphs to 12 points.
- (void) reset:(Word4046)x;  // Removes paragraph formatting that differs from the underlying style. For example, if you manually right align a paragraph and the underlying style has a different alignment, the reset method changes the alignment to match the style formatting.
- (void) space1:(Word4047)x;  // Single-spaces the specified paragraphs. The exact spacing is determined by the font size of the largest characters in each paragraph.
- (void) space15:(Word4048)x;  // Formats the specified paragraphs with 1.5-line spacing. The exact spacing is determined by adding 6 points to the font size of the largest character in each paragraph.
- (void) space2:(Word4049)x;  // Double-spaces the specified paragraphs. The exact spacing is determined by adding 12 points to the font size of the largest character in each paragraph.
- (void) tabHangingIndent:(Word4044)x count:(NSInteger)count;  // Sets a hanging indent to a specified number of tab stops. Can be used to remove tab stops from a hanging indent if the value of the count argument is a negative number.
- (void) tabIndent:(Word4045)x count:(NSInteger)count;  // Sets the left indent for the specified paragraphs to a specified number of tab stops. Can also be used to remove the indent if the value of the count argument is a negative number.
- (void) autoFit:(Word4060)x;  // Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.
- (WordTextRange *) convertRowToText:(Word4059)x separator:(WordE177)separator nestedTables:(BOOL)nestedTables;  // Converts a row to text and returns a text range object that represents the delimited text.
- (void) setLeftIndent:(Word4058)x leftIndent:(double)leftIndent rulerStyle:(WordE141)rulerStyle;  // Sets the indentation for a row or rows in a table.
- (void) setTableItemHeight:(Word4057)x rowHeight:(NSInteger)rowHeight heightRule:(WordE133)heightRule;  // Sets the height of table rows.
- (void) setTableItemWidth:(Word4056)x columnWidth:(double)columnWidth rulerStyle:(WordE141)rulerStyle;  // Sets the width of columns or cells in a table
- (void) sortAscending:(Word4054)x;  // Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) sortDescending:(Word4055)x;  // Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) tableSort:(Word4053)x excludeHeader:(BOOL)excludeHeader fieldNumber:(NSInteger)fieldNumber sortFieldType:(WordE178)sortFieldType sortOrder:(WordE179)sortOrder fieldNumberTwo:(NSInteger)fieldNumberTwo sortFieldTypeTwo:(WordE178)sortFieldTypeTwo sortOrderTwo:(WordE179)sortOrderTwo fieldNumberThree:(NSInteger)fieldNumberThree sortFieldTypeThree:(WordE178)sortFieldTypeThree sortOrderThree:(WordE179)sortOrderThree sortColumn:(BOOL)sortColumn separator:(WordE176)separator caseSensitive:(BOOL)caseSensitive languageID:(WordE182)languageID;  // Sort a table object.  For the column object only the first field is used

@end

// Represents a single AutoText entry.
@interface WordAutoTextEntry : WordBaseObject

@property (copy) NSString *autoTextValue;  // Returns or sets the value of this auto text entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of this auto text entry.
@property (copy, readonly) NSString *styleName;  // Returns the name of the style applied to the specified auto text entry.

- (WordTextRange *) insertAutoTextEntryWhere:(WordTextRange *)where richText:(BOOL)richText;  // Inserts the auto text entry in place of the specified range. Returns a text range object that represents the auto text entry.

@end

// Represents a single bookmark.
@interface WordBookmark : WordBaseObject

@property (readonly) BOOL column;  // True if the specified bookmark is a table column.
@property (readonly) BOOL empty;  // True if the specified bookmark is empty. An empty bookmark marks a location of a collapsed selection, it doesn't mark any text.
@property NSInteger endOfBookmark;  // Returns or sets the ending character position of the bookmark.
@property (copy, readonly) NSString *name;  // The name of the bookmark object.
@property NSInteger startOfBookmark;  // Returns or sets the starting character position of the bookmark.
@property (readonly) WordE160 storyType;  // Returns the story type for the bookmark.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's referred to by the bookmark.

- (WordBookmark *) copyBookmarkName:(NSString *)name;  // Sets the bookmark specified by the name argument to the location marked by another bookmark, and returns a bookmark object.

@end

// Represents options associated with border object.
@interface WordBorderOptions : WordBaseObject

@property BOOL alwaysInFront;  // Returns or sets if page borders are displayed in front of the document text.
@property WordE272 distanceFrom;  // Returns or sets a value that indicates whether the specified page border is measured from the edge of the page or from the text it surrounds.
@property NSInteger distanceFromBottom;  // Returns or sets the space in points between the text and the bottom border.
@property NSInteger distanceFromLeft;  // Returns or sets the space in points between the text and the left border.
@property NSInteger distanceFromRight;  // Returns or sets the space in points between the right edge of the text and the right border. 
@property NSInteger distanceFromTop;  // Returns or sets the space in points between the text and the top border.
@property BOOL enableBorders;  // Returns or sets border formatting for the specified object.
@property BOOL enableFirstPageInSection;  // Returns or sets if page borders are enabled for the first page in the section.
@property BOOL enableOtherPagesInSection;  // Returns or sets if page borders are enabled for all pages in the section except for the first page. 
@property (readonly) BOOL hasHorizontal;  // Returns true if a horizontal border can be applied to the object. 
@property (readonly) BOOL hasVertical;  // Returns true if a vertical border can be applied to the object. 
@property (copy) NSColor *insideColor;  // Returns or sets the RGB color of the inside borders
@property WordE110 insideColorIndex;  // Returns or sets the color index of the inside borders. 
@property WordMCoI insideColorThemeIndex;  // Returns or sets the color of the inside borders.
@property WordE167 insideLineStyle;  // Returns or sets the inside border for the specified object.
@property WordE168 insideLineWidth;  // Returns or sets the line width of the inside border of an object.
@property BOOL joinBorders;  // Returns or sets if vertical borders at the edges of paragraphs and tables are removed so that the horizontal borders can connect to the page border.
@property (copy) NSColor *outsideColor;  // Returns or sets the RGB color of the outside borders
@property WordE110 outsideColorIndex;  // Returns or sets the color index of the outside borders. 
@property WordMCoI outsideColorThemeIndex;  // Returns or sets the color of the outside borders.
@property WordE167 outsideLineStyle;  // Returns or sets the outside border for the specified object.
@property WordE168 outsideLineWidth;  // Returns or sets the line width of the outside border of an object.
@property BOOL shadow;  // Returns or sets if the specified border is formatted as shadowed.
@property BOOL surroundFooter;  // Returns or sets if a page border encompasses the document footer.
@property BOOL surroundHeader;  // Returns or sets if a page border encompasses the document header.

- (void) applyPageBordersToAllSections;  // Applies the specified page-border formatting to all sections in a document.

@end

// Represents a border of an object.
@interface WordBorder : WordBaseObject

@property WordE271 artStyle;  // Returns or sets the graphical page-border design for a document.
@property NSInteger artWidth;  // Returns or sets the width in points of the specified graphical page border.
@property (copy) NSColor *color;  // Returns or sets the RGB color for the specified border object. 
@property WordE110 colorIndex;  // Returns or sets the color index for the specified border.
@property WordMCoI colorThemeIndex;  // Returns or sets the color for the specified border or font object.
@property (readonly) BOOL inside;  // Returns true if an inside border can be applied to the specified object.
@property WordE167 lineStyle;  // Returns or sets the border line style for the specified object.
@property WordE168 lineWidth;  // Returns or sets the line width of an object's border.
@property BOOL visible;  // Returns or sets if the border object is visible


@end

// Represents the browser tool used to move the insertion point to objects in a document. This tool is comprised of the three buttons at the bottom of the vertical scroll bar.
@interface WordBrowser : WordBaseObject

@property WordE206 browserTarget;  // Returns or sets the document item that the previous and next methods locate.

- (void) nextForBrowser;  // Moves the selection to the next item indicated by the browser target. Use the browser target property to change the browser target.
- (void) previousForBrowser;  // Moves the selection to the previous item indicated by the browser target. Use the browser target property to change the browser target.

@end

@interface WordBuildingBlockCategory : WordBaseObject

- (SBElementArray *) buildingBlocks;

@property (copy, readonly) WordBuildingBlockType *buildingBlockType;  // Returns a building block type object that represents the type of building block for a building block category.
@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *name;  // Returns the name of the specified object.

- (WordBuildingBlock *) addBuildingBlockToCategoryName:(NSString *)name fromLocation:(WordTextRange *)fromLocation objectDescription:(NSString *)objectDescription insertOptions:(WordE330)insertOptions;  // Creates a new building block.

@end

// Represents a type of building block.
@interface WordBuildingBlockType : WordBaseObject

- (SBElementArray *) buildingBlockCategories;

@property (readonly) NSInteger entry_index;  // Returns an integer that represents the position of an item in a collection.
@property (copy, readonly) NSString *name;  // Returns a String that represents the localized name of a building block type.


@end

@interface WordBuildingBlock : WordBaseObject

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
@interface WordCaptionLabel : WordBaseObject

@property (readonly) BOOL builtIn;  // Returns true if this is a built-in caption label.
@property (readonly) WordE210 captionLabelId;  // Returns a constant that represents the type for the specified caption label if the built in property of the caption label object is true.
@property WordE117 captionLabelPosition;  // Returns or sets the position of caption label text.
@property NSInteger chapterStyleLevel;  // Returns or sets the heading style that marks a new chapter when chapter numbers are included with the specified caption label. The number 1 corresponds to Heading 1, 2 corresponds to Heading 2, and so on.
@property BOOL includeChapterNumber;  // Returns or sets if a chapter number is included with page numbers or a caption label.
@property (copy, readonly) NSString *name;  // Returns the name for the caption label
@property WordE153 numberStyle;  // Returns or sets the number style for the caption label object.
@property WordE120 separator;  // Returns or sets the character between the chapter number and the sequence number.


@end

// Represents a single check box form field.
@interface WordCheckBox : WordBaseObject

@property BOOL autoSize;  // True sizes the check box according to the font size of the surrounding text. False sizes the check box according to the size property.
@property BOOL checkBoxDefault;  // Returns or sets the default check box value. True if the default value is checked. 
@property BOOL checkBoxValue;  // Returns or sets if the check box is selected.  True if the check box is selected.
@property double checkboxSize;  // Returns or sets the size of the specified check box in points.
@property (readonly) BOOL valid;  // Returns if the check box object is valid.


@end

// Represents a co-authoring lock
@interface WordCoauthLock : WordBaseObject

@property BOOL isHeader_footer_lock;  // returns if this lock is a header/footer lock.
@property (copy, readonly) WordCoauthor *lockOwner;  // Get the owner of the co-authoring lock.
@property WordE332 lockType;  // Get the type of the co-authoring lock.
@property (copy, readonly) WordTextRange *textObject;  // Gets a text range object that represents the portion of a document that contains the co-authoring lock.

- (void) unlock;  // Remove the co-authoring lock.

@end

// Represents a co-authoring update
@interface WordCoauthUpdate : WordBaseObject

@property (copy, readonly) WordTextRange *textObject;  // Gets a text range object that represents the portion of a document that contains the co-authoring update.


@end

// Represents a coauthor
@interface WordCoauthor : WordBaseObject

- (SBElementArray *) coauthLocks;

@property (copy, readonly) NSString *coauthorEmailAddress;  // The email address of coauthor.
@property (copy, readonly) NSString *coauthorId;  // The ID of coauthor.
@property (copy, readonly) NSString *coauthorName;  // The name of coauthor.
@property BOOL isMe;  // returns if this coauthor is me.


@end

// Represents a coauthoring
@interface WordCoauthoring : WordBaseObject

- (SBElementArray *) coauthors;
- (SBElementArray *) coauthLocks;
- (SBElementArray *) coauthUpdates;
- (SBElementArray *) conflicts;

@property BOOL canMerge;  // returns if coauthoring can merge.
@property BOOL canShare;  // returns if coauthoring can share.
@property (copy) WordCoauthor *myself;  // returns me as author.
@property BOOL pendingUpdates;  // returns if any updates are pending.


@end

// Represents a conflict
@interface WordConflict : WordBaseObject

@property (readonly) WordE216 conflictType;  // Returns the revision type.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the conflict


@end

// Represents a custom mailing label.
@interface WordCustomLabel : WordBaseObject

@property (readonly) BOOL dotMatrix;  // True if the printer type for the specified custom label is dot matrix. False if the printer type is either laser or ink jet.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property double height;  // Returns or sets the height of the object.
@property double horizontalPitch;  // Returns or sets the horizontal distance in points between the left edge of one custom mailing label and the left edge of the next mailing label.
@property (copy) NSString *name;  // Returns or set the name of the custom mailing label.
@property NSInteger numberAcross;  // Returns or sets the number of custom mailing labels across a page.
@property NSInteger numberDown;  // Returns or sets the number of custom mailing labels down the length of a page.
@property WordE233 pageSize;  // Returns or sets the page size for the specified custom mailing label.
@property double sideMargin;  // Returns or sets the side margin widths in points for the specified custom mailing label.
@property double topMargin;  // Returns or sets the distance in points between the top edge of the page and the top boundary of the body text.
@property (readonly) BOOL valid;  // True if the various properties for example, height, width, and number down for the specified custom label work together to produce a valid mailing label.
@property double verticalPitch;  // Returns or sets the vertical distance between the top of one mailing label and the top of the next mailing label.
@property double width;  // Returns or sets the width of the object.


@end

// Represents a single data merge field in a data source.
@interface WordDataMergeDataField : WordBaseObject

@property (copy, readonly) NSString *dataMergeDataFieldValue;  // Returns the contents of the mail merge data field or mapped data field for the current record. Use the active record property to set the active record in a data merge data source.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of the data merge data field object.


@end

// Represents the data merge data source in a data merge operation.
@interface WordDataMergeDataSource : WordBaseObject

- (SBElementArray *) dataMergeFieldNames;
- (SBElementArray *) dataMergeDataFields;

@property WordE193 activeRecord;  // Returns back the index of the current active record or an enumeration specifying the record
@property (copy, readonly) NSString *connectString;  // Returns the connection string for the specified data merge data source.
@property NSInteger firstRecord;  // Returns or sets the number of the first data record to be merged in a data merge operation. 
@property (copy, readonly) NSString *headerSourceName;  // Returns the path and file name of the header source attached to the specified mail merge main document.
@property (readonly) WordE195 headerSourceType;  // Returns a value that indicates the way the header source is being supplied for the mail merge operation.
@property NSInteger lastRecord;  // Returns or sets the number of the last data record to be merged in a data merge operation. 
@property (readonly) WordE195 mailMergeDataSourceType;  // Returns the type of data merge data source.
@property (copy, readonly) NSString *name;  // Returns the name of the data merge data source.
@property (copy) NSString *queryString;  // Returns or sets the query string, SQL statement, used to retrieve a subset of the data in a data merge data source.

- (BOOL) findRecordFindText:(NSString *)findText fieldName:(NSString *)fieldName;  // Searches the contents of the specified data merge data source for text in a particular field. Returns true if the search text is found.

@end

// Represents a data merge field name in a data source.
@interface WordDataMergeFieldName : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the data merge field name object


@end

// Represents a single data merge field in a document.
@interface WordDataMergeField : WordBaseObject

@property (copy) WordTextRange *dataMergeFieldRange;  // Returns or sets a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters.
@property (readonly) WordE183 formFieldType;  // The type of this data merge field
@property BOOL locked;  // Returns or sets if the specified field is locked. When a field is locked, you cannot update the field results.
@property (copy, readonly) WordDataMergeField *nextDataMergeField;  // Returns the next data merge field
@property (copy, readonly) WordDataMergeField *previousMakeMergeField;  // Returns the previous data merge field


@end

// Represents the data merge functionality in Word.
@interface WordDataMerge : WordBaseObject

- (SBElementArray *) dataMergeFields;

@property (copy, readonly) WordDataMergeDataSource *dataSource;  // Returns or sets the destination of the data merge results.
@property WordE192 destination;  // Returns or sets the destination of the data merge results.
@property (copy) NSString *mailAddressFieldName;  // Returns or sets the name of the field that contains e-mail addresses that are used when the data merge destination is electronic mail.
@property BOOL mailAsAttachment;  // Returns or sets if the merge documents are sent as attachments when the data merge destination is an e-mail message or a fax.
@property (copy) NSString *mailSubject;  // Returns or sets the subject line used when the data merge destination is electronic mail.
@property WordE190 mainDocumentType;  // Returns or sets the data merge main document type.
@property (readonly) WordE191 state;  // Returns the current state of a data merge operation.
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
- (WordDataMergeField *) makeNewDataMergeIfFieldTextRange:(WordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(WordE196)comparison compareTo:(NSString *)compareTo trueText:(NSString *)trueText falseText:(NSString *)falseText;  // Create a new data merge if field
- (WordDataMergeField *) makeNewDataMergeNextFieldTextRange:(WordTextRange *)textRange;  // Create a new data merge next field
- (WordDataMergeField *) makeNewDataMergeNextIfFieldTextRange:(WordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(WordE196)comparison compareTo:(NSString *)compareTo;  // Create a new data merge next if field
- (WordDataMergeField *) makeNewDataMergeRecFieldTextRange:(WordTextRange *)textRange;  // Create a new data merge rec field
- (WordDataMergeField *) makeNewDataMergeSequenceFieldTextRange:(WordTextRange *)textRange;  // Create a new data merge sequence field
- (WordDataMergeField *) makeNewDataMergeSetFieldTextRange:(WordTextRange *)textRange name:(NSString *)name valueText:(NSString *)valueText;  // Create a new data merge set field
- (WordDataMergeField *) makeNewDataMergeSkipIfFieldTextRange:(WordTextRange *)textRange mergeField:(NSString *)mergeField comparison:(WordE196)comparison compareTo:(NSString *)compareTo;  // Create a new data merge skip if field
- (void) openDataSourceName:(NSString *)name format:(WordE162)format confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly linkToSource:(BOOL)linkToSource addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate connection:(NSString *)connection SQLStatement:(NSString *)SQLStatement SQLStatement1:(NSString *)SQLStatement1;  // Attaches a data source to the specified document, which becomes a main document if it's not one already.
- (void) openHeaderSourceName:(NSString *)name format:(WordE162)format confirmConversions:(BOOL)confirmConversions readOnly:(BOOL)readOnly addToRecentFiles:(BOOL)addToRecentFiles passwordDocument:(NSString *)passwordDocument passwordTemplate:(NSString *)passwordTemplate revert:(BOOL)revert writePassword:(NSString *)writePassword writePasswordTemplate:(NSString *)writePasswordTemplate;  // Attaches a mail merge header source to the specified document.
- (void) useAddressBookBookType:(NSString *)bookType;  // Selects the address book that's used as the data source for a mail merge operation.

@end

// Contains global application-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page. You can return or set attributes either at the application global level or at the document level.
@interface WordDefaultWebOptions : WordBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page.
@property BOOL alwaysSaveInDefaultEncoding;  // Returns or saves if the default encoding is used when you save a Web page or plain text document, independent of the file's original encoding when opened.  The default value is False.
@property BOOL checkIfOfficeIsHtmleditor;  // Returns or sets if Microsoft Word checks to see whether an Office application is the default HTML editor when you start Word.
@property BOOL checkIfWordIsDefaultHtmleditor;  // Returns or sets if Microsoft Word checks to see whether it is the default HTML editor when you start Word. The default value is true.
@property WordMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property WordMSsz screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL updateLinksOnSave;  // Returns or sets if hyperlinks and paths to all supporting files are automatically updated before you save the document as a Web page, ensuring that the links are up-to-date at the time the document is saved.
@property BOOL useLongFileNames;  // Returns or sets if long file names are used when you save the document as a Web page.


@end

// Represents a built-in dialog box.
@interface WordDialog : WordBaseObject

@property WordE185 defaultDialogTab;  // Returns or sets the active tab when the specified dialog box is displayed.
@property (readonly) WordE186 dialogType;  // The built-in dialog this object represents.

- (NSInteger) displayWordDialogTimeOut:(NSInteger)timeOut;  // Displays the specified built-in Word dialog box until either the user closes it or the specified amount of time has passed. Returns a Long that indicates which button was clicked to close the dialog box.
- (void) executeDialog;  // Applies the current settings of a Microsoft Word dialog box.
- (NSInteger) showTimeOut:(NSInteger)timeOut;  // Displays and carries out actions initiated in the specified built-in Word dialog box. Returns a number which indicates the button used to dismiss the dialog box.

@end

// Represents a single version of a document.
@interface WordDocumentVersion : WordBaseObject

@property (copy, readonly) NSString *comment;  // Returns the comment associated with the specified version of a document.
@property (copy, readonly) NSDate *dateValue;  // The date and time that the document version was saved.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *savedBy;  // Returns the name of the user who saved the specified version of the document.

- (void) openVersion;  // Opens the specified version of the document.

@end

// Represents a Microsoft Word document.
@interface WordDocument : WordBaseObject

- (SBElementArray *) documentProperties;
- (SBElementArray *) customDocumentProperties;
- (SBElementArray *) bookmarks;
- (SBElementArray *) tables;
- (SBElementArray *) footnotes;
- (SBElementArray *) endnotes;
- (SBElementArray *) WordComments;
- (SBElementArray *) sections;
- (SBElementArray *) paragraphs;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) characters;
- (SBElementArray *) fields;
- (SBElementArray *) formFields;
- (SBElementArray *) WordStyles;
- (SBElementArray *) frames;
- (SBElementArray *) tablesOfFigures;
- (SBElementArray *) variables;
- (SBElementArray *) revisions;
- (SBElementArray *) tablesOfContents;
- (SBElementArray *) tablesOfAuthorities;
- (SBElementArray *) windows;
- (SBElementArray *) indexes;
- (SBElementArray *) subdocuments;
- (SBElementArray *) storyRanges;
- (SBElementArray *) hyperlinkObjects;
- (SBElementArray *) shapes;
- (SBElementArray *) listTemplates;
- (SBElementArray *) WordLists;
- (SBElementArray *) inlineShapes;
- (SBElementArray *) documentVersions;
- (SBElementArray *) readabilityStatistics;
- (SBElementArray *) grammaticalErrors;
- (SBElementArray *) spellingErrors;
- (SBElementArray *) mathObjects;

@property (copy, readonly) WordWindow *activeWindow;  // Returns a window object that represents the active window, the window with the focus. If there are no windows open, an error occurs.
@property (copy) WordTemplate *attachedTemplate;  // Returns of sets a template object that represents the template attached to the specified document. To set this property, specify either the name of the template or a template object.
@property BOOL autoHyphenation;  // Returns or sets if automatic hyphenation is turned on for the specified document.
@property (copy) WordShape *backgroundShape;  // Returns or sets a shape object that represents the background image for the specified document.
@property WordE184 clickAndTypeParagraphStyle;  // Returns or sets the default paragraph style applied to text by the Click and Type feature in the specified document.
@property (copy, readonly) WordCoauthoring *coauthoring;
@property WordE314 compatibleVersion;  // Returns or sets the compatibility options to a given application version.
@property NSInteger consecutiveHyphensLimit;  // Returns or sets the maximum number of consecutive lines that can end with hyphens. If this property is set to zero, any number of consecutive lines can end with hyphens.
@property (copy, readonly) WordDataMerge *dataMerge;  // Returns the data merge object.
@property double defaultTabStop;  // Returns or sets the interval in points between the default tab stops in the specified document.
@property (readonly) WordE184 defaultTableStyle;  // Returns the default table style for this document.
@property (copy, readonly) WordOfficeTheme *documentTheme;  // Returns an office theme object that represents the Microsoft Office theme applied to a document.
@property (readonly) WordE222 document_type;  // Returns the document type.
@property WordE108 eastAsianLineBreak;  // Returns or sets the line break control level for the specified document. This property is ignored if the “east asian line break control” property is set to false. Note that “east asian line break control” is a paragraph format property.
@property BOOL embedTrueTypeFonts;  // Returns or set if Microsoft Word embeds TrueType fonts in a document when it's saved. This allow others to view the document with the same fonts that were used to create it.
@property (copy, readonly) WordEndnoteOptions *endnoteOptions;  // Returns the endnote options object.
@property (copy, readonly) WordEnvelope *envelopeObject;  // Returns the envelop object.
@property (copy, readonly) WordFieldOptions *fieldOptions;
@property (copy, readonly) WordFootnoteOptions *footnoteOptions;  // Returns the footnote options object.
@property (copy, readonly) NSString *fullName;  // Returns the full name of the document.
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
@property WordE107 justificationMode;  // Returns or sets the character spacing adjustment for the specified document.
@property (copy) WordLetterContent *letterContent;  // Return or sets the letter content object associated with the document.
@property (readonly) WordE327 mathBinaryOperatorBreak;  // Gets or sets a value that specifies where Microsoft Word places binary operators when equations span two or more lines.
@property (readonly) WordE326 mathDefaultJustification;  // Gets or sets a value that indicates the default justification.
@property (copy, readonly) NSString *mathFontName;  // Gets or sets the name of the font that is used in a document to display equations.
@property (readonly) double mathLeftMargin;  // Gets or sets a value that specifies the left margin for equations.
@property (readonly) double mathRightMargin;  // Gets or sets a value that specifies the right margin for equations.
@property (readonly) WordE328 mathSubtractionOperator;  // Gets or sets a value that specifies how Microsoft Word handles a subtraction operator that falls before a line break.
@property (readonly) double mathWrap;  // Gets or sets a value that specifies the placement of the second line of an equation that wraps to a new line.
@property (copy, readonly) NSString *name;  // Returns the name of the document.
@property (readonly) BOOL naryUsesSubscripts;  // Gets or sets a value that specifies the default location of limits for n-ary objects other than integrals
@property (copy) NSString *noLineBreakAfter;  // Returns or sets the kinsoku characters after which Microsoft Word will not break a line.
@property (copy) NSString *noLineBreakBefore;  // Returns or sets the kinsoku characters before which Microsoft Word will not break a line.
@property (copy) WordPageSetup *pageSetup;  // Returns or sets the page setup object.
@property (copy) NSString *password;  // Sets a password that must be supplied to open the specified document. This is write-only property
@property (copy, readonly) NSString *path;  // Returns the path to the document.
@property BOOL printFormsData;  // Returns or sets if Microsoft Word prints onto a preprinted form only the data entered in the corresponding online form.
@property BOOL printPostScriptOverText;  // Returns or sets if PRINT field instructions such as PostScript commands in a document are to be printed on top of text and graphics when a PostScript printer is used.
@property BOOL printRevisions;  // Returns or sets if revision marks are printed with the document. False if revision marks aren't printed that is, tracked changes are printed as if they'd been accepted.
@property (readonly) WordE234 protectionType;  // Returns the protection type for the specified document.
@property (readonly) BOOL readOnly;  // True if changes to the document cannot be saved to the original document.
@property BOOL readOnlyRecommended;  // Returns or set if Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only.
@property BOOL removeDateAndTime;  // Returns or sets if Microsoft Word removes date and time from revisions upon saving a document.
@property BOOL removePersonalInformation;  // Returns or sets if Microsoft Word removes all user information from comments, revisions, and the properties dialog box upon saving a document.
@property (readonly) WordE161 saveFormat;  // Returns the file format of the specified document or file converter. Will be a unique number that specifies an external file converter or a constant.
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
@property (copy, readonly) NSArray *unavailableFonts;  // Returns a list of fonts used in the document that are not available on the current system.
@property BOOL updateStylesOnOpen;  // Returns or sets if the styles in the specified document are updated to match the styles in the attached template each time the document is opened.
@property (readonly) BOOL useDefaultMathSettings;  // Gets or sets a value that indicates whether to use the default math settings when creating new equations.
@property (readonly) BOOL useSmallFractions;  // Gets or sets a value that indicates whether to use small fractions in equations contained within the document.
@property (copy, readonly) WordWebOptions *webOptions;  // Returns the web options object.
@property (copy) NSString *writePassword;  // Sets a password for saving changes to the specified document. This is write-only property
@property (readonly) BOOL writeReserved;  // True if the specified document is protected with a write password.

- (void) acceptAllRevisions;  // Accepts all tracked changes in the document.
- (void) acceptAllShownRevisions;  // Accepts all shown tracked changes.
- (void) applyDocumentThemeFileName:(NSString *)fileName;  // Applies an Office theme to the document.
- (void) autoFormat;  // Automatically formats a document
- (BOOL) canCheckIn;  // Returns True if Word can check in a specified document to a server.
- (void) changeDefaultTableStyleStyle:(WordE184)style changeInTemplate:(BOOL)changeInTemplate;  // Sets the default table style.
- (void) checkConsistency;  // Searches all text in a Japanese language document and displays instances where character usage is inconsistent for the same words.
- (void) checkInSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic;  // Returns a document from a local computer to a server, and sets the local document to read-only so that it cannot be edited locally.
- (void) checkInWithVersionSaveChanges:(BOOL)saveChanges comments:(NSString *)comments makePublic:(BOOL)makePublic versionType:(WordE331)versionType;  // Returns a document from a local computer to a server, and sets the local file to read-only so that it cannot be edited locally.
- (void) closePrintPreview;  // Switches the specified document from print preview to the previous view. If the specified document isn't in print preview, an error occurs.
- (void) comparePath:(NSString *)path authorName:(NSString *)authorName target:(WordE297)target detectFormatChanges:(BOOL)detectFormatChanges ignoreAllComparisonWarnings:(BOOL)ignoreAllComparisonWarnings addToRecentFiles:(BOOL)addToRecentFiles;  // Compares this document with another.
- (NSInteger) computeStatisticsStatistic:(WordE155)statistic includeFootnotesAndEndnotes:(BOOL)includeFootnotesAndEndnotes;  // Computes a set of readability statistics for this document.  This must be called before accessing the readability statistics for this document.
- (void) copyStylesFromTemplateTemplate:(NSString *)template_;  // Copies styles from the specified template to a document.
- (WordLetterContent *) createLetterContentDateFormat:(NSString *)dateFormat includeHeaderFooter:(BOOL)includeHeaderFooter pageDesign:(NSString *)pageDesign letterStyle:(WordE245)letterStyle letterhead:(BOOL)letterhead letterheadLocation:(WordE246)letterheadLocation letterheadSize:(double)letterheadSize recipientName:(NSString *)recipientName recipientAddress:(NSString *)recipientAddress salutation:(NSString *)salutation salutationType:(WordE247)salutationType recipientReference:(NSString *)recipientReference mailingInstructions:(NSString *)mailingInstructions attentionLine:(NSString *)attentionLine subject:(NSString *)subject ccList:(NSString *)ccList returnAddress:(NSString *)returnAddress senderName:(NSString *)senderName closing:(NSString *)closing senderCompany:(NSString *)senderCompany senderJobTitle:(NSString *)senderJobTitle senderInitials:(NSString *)senderInitials enclosureCount:(NSInteger)enclosureCount;  // Create a new letter content object for use with the letter wizard.
- (WordTextRange *) createRangeStart:(NSInteger)start end:(NSInteger)end;  // Returns a text range object by using the specified starting and ending character positions.
- (void) dataForm;  // Displays the data form dialog box, in which you can modify data records. You can use this method with a mail merge main document, a mail merge data source, or any document that contains data delimited by table cells or separator characters
- (void) deleteAllComments;  // Deletes all comments.
- (void) deleteAllShownComments;  // Deletes all shown comments.
- (void) fitToPages;  // Decreases the font size of text just enough so that the document will fit on one fewer pages. An error occurs if Word is unable to reduce the page count by one.
- (void) followHyperlinkAddress:(NSString *)address subAddress:(NSString *)subAddress newWindow:(BOOL)newWindow addHistory:(BOOL)addHistory extraInfo:(NSString *)extraInfo;  // This method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application. If the hyperlink uses the file protocol, this method opens the document instead of downloading it.
- (NSString *) getActiveWritingStyleLanguageID:(WordE182)languageID;  // Returns the writing style for a specified language.
- (NSArray *) getCrossReferenceItemsReferenceType:(WordE211)referenceType;  // Returns an list of items that can be cross-referenced based on the specified cross-reference type.
- (BOOL) getDocumentCompatibilityCompatibilityItem:(WordE231)compatibilityItem;  // Returns the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.
- (WordStoryRange *) getStoryRangeStoryType:(WordE160)storyType;  // Returns a range object that represents the story specified by the story type argument.
- (void) makeCompatibilityDefault;  // Uses the correct settings of the document compatibility options set by the set compatibility options method as the default for new documents.
- (void) manualHyphenation;  // Initiates manual hyphenation of a document, one line at a time. The user is prompted to accept or decline suggested hyphenations.
- (WordField *) markEntryForTableOfContentsRange:(WordTextRange *)range entry:(NSString *)entry tableID:(NSString *)tableID level:(NSInteger)level;  // Inserts a table of contents entry field after the specified range. The method returns a field object representing the new field.
- (WordField *) markEntryForTableOfFiguresRange:(WordTextRange *)range entry:(NSString *)entry tableID:(NSString *)tableID level:(NSInteger)level;  // Inserts a table of figures entry field after the specified range. The method returns a field object representing the new field.
- (WordField *) markForIndexRange:(WordTextRange *)range entry:(NSString *)entry crossReference:(NSString *)crossReference bookmarkName:(NSString *)bookmarkName;  // Inserts an index entry field after the specified range. The method returns a field object representing the new field.
- (void) mergeFileName:(NSString *)fileName;  // Merges the changes marked with revision marks from one document to another.
- (void) mergeSubdocumentsFirstSubdocument:(WordSubdocument *)firstSubdocument lastSubdocument:(WordSubdocument *)lastSubdocument;  // Merges the specified subdocuments of a master document into a single subdocument.
- (void) presentIt;  // Opens PowerPoint with the specified Word document loaded.
- (void) printPreview;  // Switches the view to print preview.
- (void) protectProtectionType:(WordE234)protectionType noReset:(BOOL)noReset password:(NSString *)password useIrm:(BOOL)useIrm enforceStyleLocks:(BOOL)enforceStyleLocks;  // Helps to protect the specified document from changes. When a document is protected, users can make only limited changes, such as adding annotations, making revisions, or completing a form.
- (BOOL) redoTimes:(NSInteger)times;  // Redoes the last action that was undone. It reverses the undo method. Returns true if the actions were redone successfully.
- (void) rejectAllRevisions;  // Rejects all tracked changes in the document.
- (void) rejectAllShownRevisions;  // Rejects all shown tracked changes.
- (void) reload;  // Reloads a cached document by resolving the hyperlink to the document and downloading it.
- (void) removeReminder;  // Remove the document's reminder.
- (void) repaginate;  // Repaginates the entire document
- (void) runLetterWizardLetterContent:(WordLetterContent *)letterContent wizardMode:(BOOL)wizardMode;  // Runs the Letter Wizard on the document
- (void) saveAsFileName:(NSString *)fileName fileFormat:(WordE161)fileFormat lockComments:(BOOL)lockComments password:(NSString *)password addToRecentFiles:(BOOL)addToRecentFiles writePassword:(NSString *)writePassword readOnlyRecommended:(BOOL)readOnlyRecommended embedTruetypeFonts:(BOOL)embedTruetypeFonts saveNativePictureFormat:(BOOL)saveNativePictureFormat saveFormsData:(BOOL)saveFormsData textEncoding:(NSInteger)textEncoding insertLineBreaks:(BOOL)insertLineBreaks allowSubstitutions:(BOOL)allowSubstitutions lineEndingType:(WordE311)lineEndingType HTMLDisplayOnlyOutput:(BOOL)HTMLDisplayOnlyOutput maintainCompatibility:(BOOL)maintainCompatibility;  // Saves the document with a new name or format.
- (void) saveVersionComment:(NSString *)comment;  // Saves a version of the document with a comment.
- (void) sendHtmlMail;  // Opens a message window for sending the specified document, formatted as html, through Microsoft Outlook.
- (void) sendMail;  // Opens a message window for sending the specified document through your registered mail program.
- (void) setActiveWritingStyleLanguageID:(WordE182)languageID writingStyle:(NSString *)writingStyle;  // Sets the writing style for the specified language.
- (void) setDocumentCompatibilityCompatibilityItem:(WordE231)compatibilityItem isCompatible:(BOOL)isCompatible;  // Sets the current state of the specified compatibility item for this document. Compatibility options affect how a document is displayed in Microsoft Word.
- (void) setReminderAt:(NSDate *)at;  // Create or modify a reminder for the document.
- (BOOL) undoTimes:(NSInteger)times;  // Undoes the last action or a sequence of actions, which are displayed in the undo list. Returns true if the actions were successfully undone.
- (void) undoClear;  // Clear the list of actions that can be undone.
- (void) unprotectPassword:(NSString *)password;  // Removes protection from the specified document. If the document isn't protected, this method generates an error.
- (void) updateStyles;  // Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name.
- (void) upgrade;  // upgrade document
- (void) viewPropertyBrowser;  // Displays the property window for the selected control in the specified document.
- (void) webPagePreview;  // Displays a preview of the document as it would look if saved as a Web page.

@end

// Represents a dropped capital letter at the beginning of a paragraph.
@interface WordDropCap : WordBaseObject

@property double distanceFromText;  // Returns or sets the distance in points between the dropped capital letter and the paragraph text. 
@property WordE172 dropPosition;  // Returns or sets the position of a dropped capital letter.
@property (copy) NSString *fontName;  // Returns or sets the name of the font for the dropped capital letter.
@property NSInteger linesToDrop;  // Returns or sets the height in lines of the specified dropped capital letter.

- (void) enable;  // Formats the first character in the specified paragraph as a dropped capital letter.

@end

// Represents a drop-down form field that contains a list of items in a form.
@interface WordDropDown : WordBaseObject

- (SBElementArray *) listEntries;

@property NSInteger dropDownDefault;  // Returns or sets the default drop-down item. The first item in a drop-down form field is 1, the second item is 2, and so on.
@property NSInteger dropDownValue;  // Returns or sets the number of the selected item in a drop-down form field.
@property (readonly) BOOL valid;  // Returns if the drop down object is valid.


@end

// A representation of the options associated with endnotes.
@interface WordEndnoteOptions : WordBaseObject

@property (copy, readonly) WordTextRange *endnoteContinuationNotice;  // Returns a text range object that represents the endnote continuation notice.
@property (copy, readonly) WordTextRange *endnoteContinuationSeparator;  // Returns a text range object that represents the endnote continuation separator.
@property WordE175 endnoteLocation;  // Returns or sets the position of all endnotes.
@property WordE152 endnoteNumberStyle;  // Returns or sets the number style for endnotes.
@property WordE173 endnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks.
@property (copy, readonly) WordTextRange *endnoteSeparator;  // Returns a text range object that represents the endnote separator.
@property NSInteger endnoteStartingNumber;  // Returns or sets the starting endnote number.

- (void) endnoteConvert;  // Converts endnotes to footnotes, or vice versa.
- (void) swapWithFootnotes;  // Converts all footnotes in a document to endnotes and vice versa.

@end

// Represents an endnote.
@interface WordEndnote : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *noteReference;  // Returns a text range object that represents a endnote mark.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the endnote object.


@end

// Represents an envelope.
@interface WordEnvelope : WordBaseObject

@property (copy, readonly) WordTextRange *address;  // Returns the envelope delivery address as a text range object.
@property double addressFromLeft;  // Returns or sets the distance in points between the left edge of the envelope and the delivery address.
@property double addressFromTop;  // Returns or sets the distance in points between the top edge of the envelope and the delivery address.
@property (copy, readonly) WordWordStyle *addressStyle;  // Returns a Word style object that represents the delivery address style for the envelope.
@property BOOL defaultFaceUp;  // Returns or sets if envelopes are fed face up by default.
@property double defaultHeight;  // Returns or sets the default envelope height, in points.
@property BOOL defaultOmitReturnAddress;  // Returns or sets if the return address is omitted from envelopes by default.
@property WordE244 defaultOrientation;  // Returns or sets the default orientation for feeding envelopes.
@property BOOL defaultPrintFIMA;  // Returns or sets if a Facing Identification Mark FIM-A to envelopes by default. A FIM-A code is used to presort courtesy reply mail. The default print barcode property must be set to true before this property is set.
@property BOOL defaultPrintBarCode;  // Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set
@property (copy) NSString *defaultSize;  // Returns or sets the default envelope size. If you set either the default height or default width property, the property  is automatically changed to return custom size.
@property double defaultWidth;  // Returns or sets the default envelope width, in points.
@property WordE207 feedSource;  // Returns or sets the paper tray for the envelope.
@property (copy, readonly) WordTextRange *returnAddress;  // Returns the envelope return address as a text range object.
@property double returnAddressFromLeft;  // Returns or sets the distance in points between the left edge of the envelope and the return address.
@property double returnAddressFromTop;  // Returns or sets the distance in points between the top edge of the envelope and the return address.
@property (copy, readonly) WordWordStyle *returnAddressStyle;  // Returns a Word style object that represents the return address style for the envelope.

- (void) insertEnvelopeDataExtractAddress:(BOOL)extractAddress address:(NSString *)address autoText:(NSString *)autoText omitReturnAddress:(BOOL)omitReturnAddress returnAddress:(NSString *)returnAddress returnAutotext:(NSString *)returnAutotext printBarCode:(BOOL)printBarCode printFIMA:(BOOL)printFIMA envelopeSize:(NSString *)envelopeSize envelopeHeight:(NSInteger)envelopeHeight envlopeWidth:(NSInteger)envlopeWidth feedSource:(BOOL)feedSource addressFromLeft:(NSInteger)addressFromLeft addressFromTop:(NSInteger)addressFromTop returnAddressFromLeft:(NSInteger)returnAddressFromLeft returnAddressFromTop:(NSInteger)returnAddressFromTop defaultFaceUp:(BOOL)defaultFaceUp defaultOrientation:(WordE244)defaultOrientation sizeFromPageSetup:(BOOL)sizeFromPageSetup showPageSetupDialog:(BOOL)showPageSetupDialog createNewDocument:(BOOL)createNewDocument;  // Inserts an envelope as a separate section at the beginning of the specified document.
- (void) printOutEnvelopeExtractAddress:(BOOL)extractAddress address:(NSString *)address autoText:(NSString *)autoText omitReturnAddress:(BOOL)omitReturnAddress returnAddress:(NSString *)returnAddress returnAutotext:(NSString *)returnAutotext printBarCode:(BOOL)printBarCode printFIMA:(BOOL)printFIMA envelopeSize:(NSString *)envelopeSize envelopeHeight:(NSInteger)envelopeHeight envlopeWidth:(NSInteger)envlopeWidth feedSource:(BOOL)feedSource addressFromLeft:(NSInteger)addressFromLeft addressFromTop:(NSInteger)addressFromTop returnAddressFromLeft:(NSInteger)returnAddressFromLeft returnAddressFromTop:(NSInteger)returnAddressFromTop defaultFaceUp:(BOOL)defaultFaceUp defaultOrientation:(WordE244)defaultOrientation sizeFromPageSetup:(BOOL)sizeFromPageSetup showPageSetupDialog:(BOOL)showPageSetupDialog;
- (void) updateDocument;  // Updates the envelope in the document with the current envelope settings.

@end

@interface WordFieldOptions : WordBaseObject

@property BOOL locked;


@end

// Represents a field.
@interface WordField : WordBaseObject

@property (readonly) NSInteger entryIndex;  // HACK: Modification for Word 2004 support.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) WordTextRange *fieldCode;  // Returns a text range object that represents a field's code. A field's code is everything that's enclosed by the field characters including the leading space and trailing space characters.
@property (readonly) WordE187 fieldKind;  // Returns the type of link for a field object.
@property (copy) NSString *fieldText;  // Returns or sets data in an ADDIN field. The data is not visible in the field code or result. It is only accessible by returning the value of the data property. If the field isn't an ADDIN field, this property will cause an error.
@property (readonly) WordE183 fieldType;  // Returns the field type.
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
@interface WordFileConverter : WordBaseObject

@property (readonly) BOOL canOpen;  // Returns true if the specified file converter is designed to open files.
@property (readonly) BOOL canSave;  // Returns true if the specified file converter is designed to save files.
@property (copy, readonly) NSString *className;  // Returns a unique name that identifies the file converter.
@property (copy, readonly) NSString *extensions;  // Returns the file name extensions associated with the specified file converter object.
@property (copy, readonly) NSString *formatName;  // Returns the name of the specified file converter.
@property (copy, readonly) NSString *name;  // Returns the name of the file converter.
@property (readonly) NSInteger openFormat;  // Returns the file format of the specified file converter. It will be a unique number that represents an external file converter.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified file converter. 
@property (readonly) NSInteger saveFormat;  // Returns the file format of the specified document or file converter. It will be a unique number that specifies an external file converter.


@end

// Represents the criteria for a find operation.
@interface WordFind : WordBaseObject

@property (copy) NSString *content;  // Returns or sets the text in the find object.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this find object.
@property BOOL format;  // Returns or set if formatting is included in the find operation.
@property BOOL forward;  // Returns or sets if the find operation searches forward through the document. False if it searches backward through the document.
@property (readonly) BOOL found;  // True if the search produces a match.
@property (copy, readonly) WordFrame *frame;  // Returns the frame object associated with the find object.
@property NSInteger highlight;  // Returns or sets if highlight formatting is included in the find criteria
@property WordE182 languageID;  // Returns or sets the language for the find object
@property WordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
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
@property WordE184 style;  // Returns or sets the Word style associated with the find object.
@property WordE182 supplementalLanguageID;  // Returns or sets the language for the text range object
@property WordE265 wrap;  // Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached or vice versa if forward is set to false or if the search text isn't found in the specified selection or range. 

- (void) clearAllFuzzyOptions;  // Clears all nonspecific search options associated with Japanese text.
- (WordEFRt) executeFindFindText:(NSString *)findText matchCase:(BOOL)matchCase matchWholeWord:(BOOL)matchWholeWord matchWildcards:(BOOL)matchWildcards matchSoundsLike:(BOOL)matchSoundsLike matchAllWordForms:(BOOL)matchAllWordForms matchForward:(BOOL)matchForward wrapFind:(WordE265)wrapFind findFormat:(BOOL)findFormat replaceWith:(NSString *)replaceWith replace:(WordE273)replace;  // Runs the specified find operation. Returns true if the find operation is successful.
- (void) setAllFuzzyOptions;  // Activates all nonspecific search options associated with Japanese text.

@end

// Contains font attributes, such as font name, size, and color, for an object.
@interface WordFont : WordBaseObject

@property BOOL allCaps;  // Returns or sets if the font is formatted as all capital letters.
@property WordE124 animation;  // Returns or sets the type of animation applied to the font.
@property (copy) NSString *asciiName;  // Returns or sets the font used for Latin text characters with character codes from 0 through 127. 
@property BOOL bold;  // Returns or sets if the font is formatted as bold.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property (copy) NSColor *color;  // Returns or sets the RGB color of the font object.
@property WordE110 colorIndex;  // Returns or sets the color for the font object using an index.
@property WordMCoI colorThemeIndex;  // Returns or sets the color for the specified border or font object.
@property BOOL disableCharacterSpaceGrid;  // Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding font object.
@property BOOL doubleStrikeThrough;  // Returns or sets if the specified font is formatted as double strikethrough text.
@property (copy) NSString *eastAsianName;  // Returns or sets an East Asian font name.
@property BOOL emboss;  // Returns or sets if the specified font is formatted as embossed.
@property WordE114 emphasisMark;  // Returns or sets the emphasis mark for a character or designated character string.
@property BOOL engrave;  // Returns or sets if the specified font is formatted as engraved.
@property NSInteger fontPosition;  // Returns or sets the position of text in points relative to the base line. A positive number raises the text, and a negative number lowers it.
@property double fontSize;  // Returns or sets the font size.
@property BOOL hidden;  // Returns or sets if the font is formatted as hidden text.
@property BOOL italic;  // Returns or sets if the font is formatted as italic.
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
@property WordE113 underline;  // Returns or sets the type of underline applied to the font.
@property (copy) NSColor *underlineColor;  // Returns or sets the RGB color of the underline for the font object.
@property WordMCoI underlineColorThemeIndex;  // Returns a value specifying the color of the underline for the selected text.

- (void) growFont;  // Increases the font size to the next available size. If the selection or range contains more than one font size, each size is increased to the next available setting.
- (void) reset;  // Removes changes that were made to a font.
- (void) setAsFontTemplateDefault;  // Sets the font formatting as the default for the active document and all new documents based on the active template. The default font formatting is stored in the Normal style.
- (void) shrinkFont;  // Decreases the font size to the next available size. If the selection or range contains more than one font size, each size is decreased to the next available setting.

@end

// A representation of the options associated with footnotes.
@interface WordFootnoteOptions : WordBaseObject

@property (copy, readonly) WordTextRange *footnoteContinuationNotice;  // Returns a text range object that represents the footnote continuation notice.
@property (copy, readonly) WordTextRange *footnoteContinuationSeparator;  // Returns a text range object that represents the footnote continuation separator.
@property WordE174 footnoteLocation;  // Returns or sets the position of all footnotes.
@property WordE152 footnoteNumberStyle;  // Returns or sets the number style for footnotes.
@property WordE173 footnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. 
@property (copy, readonly) WordTextRange *footnoteSeparator;  // Returns a text range object that represents the footnote separator.
@property NSInteger footnoteStartingNumber;  // Returns or sets the starting footnote number. 

- (void) footnoteConvert;  // Converts endnotes to footnotes, or vice versa.
- (void) swapWithEndnotes;  // Converts all footnotes in a document to endnotes and vice versa.

@end

// Represents a footnote positioned at the bottom of the page or beneath text.
@interface WordFootnote : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *noteReference;  // Returns a text range object that represents a footnote mark.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the footnote object.


@end

// Represents a single form field.
@interface WordFormField : WordBaseObject

@property BOOL calculateOnExit;  // Returns or sets if references to the specified form field are automatically updated whenever the field is exited.
@property (copy, readonly) WordCheckBox *checkBox;  // Returns the check box object associated with the form filed object.
@property (copy, readonly) WordDropDown *dropDown;  // Returns the drop down object associated with the form filed object.
@property BOOL enabled;  // Returns or sets if this form field object is enabled.
@property (copy) NSString *formFieldResult;  // Returns or sets a string that represents the result of the specified form field.
@property (readonly) WordE183 formFieldType;  // The type of this form field.
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
@interface WordFrame : WordBaseObject

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this frame object.
@property double height;  // Returns or sets the height of the object.
@property WordE134 heightRule;  // Returns or sets the rule for determining the height of the specified frame.
@property double horizontalDistanceFromText;  // Returns or sets the horizontal distance between a frame and the surrounding text, in points.
@property double horizontalPosition;  // Returns or sets the horizontal distance between the edge of the frame and the item specified by the relative horizontal position property.
@property BOOL lockAnchor;  // Returns or sets if the specified frame is locked. The frame anchor indicates where the frame will appear in Draft view. You cannot reposition a locked frame anchor.
@property WordE236 relativeHorizontalPosition;  // Returns or sets what the horizontal position of a frame is relative.
@property WordE237 relativeVerticalPosition;  // Returns or sets what the vertical position of a frame is relative.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the frame object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the frame object.
@property BOOL textWrap;  // Returns or sets if document text wraps around the specified frame.
@property double verticalDistanceFromText;  // Returns or sets the vertical distance in points between a frame and the surrounding text.
@property double verticalPosition;  // Returns or sets the vertical distance between the edge of the frame and the item specified by the relative vertical position property. 
@property double width;  // Returns or sets the width of the object.
@property WordE134 widthRule;  // Returns or sets the rule used to determine the width of a frame.


@end

// Represents a single header or footer.
@interface WordHeaderFooter : WordBaseObject

- (SBElementArray *) pageNumbers;
- (SBElementArray *) shapes;

@property (readonly) WordE163 headerFooterIndex;  // Returns a constant that represents the specified header or footer in a document or section.
@property (readonly) BOOL isHeader;  // Returns true if this object is a header.
@property BOOL linkToPrevious;  // Returns or sets if the specified header or footer is linked to the corresponding header or footer in the previous section. When a header or footer is linked, its contents are the same as in the previous header or footer.
@property (copy, readonly) WordPageNumberOptions *pageNumberOptions;  // Return the page number options object associated with this header footer object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the header or footer.


@end

// Represents a style used to build a table of contents or figures.
@interface WordHeadingStyle : WordBaseObject

@property NSInteger level;  // Returns or sets the level for the heading style in a table of contents or table of figures.
@property WordE184 style;  // Returns or sets the style associated with the heading style object.


@end

// Represents a hyperlink.
@interface WordHyperlinkObject : WordBaseObject

@property (copy) NSString *emailSubject;  // Returns or sets the text string for the specified hyperlink’s subject line. The subject line is appended to the hyperlink’s Internet address, or URL.
@property (readonly) BOOL extraInfoRequired;  // Returns true if extra information is required to resolve the specified hyperlink.
@property (copy) NSString *hyperlinkAddress;  // Returns or sets the address, for example, a file name or URL of the specified hyperlink. 
@property (readonly) WordMHlT hyperlinkType;  // The type of this hyperlink
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
@interface WordIndex : WordBaseObject

@property BOOL accentedLetters;  // Returns or sets if the specified index contains separate headings for accented letters, for example, words that begin with À are under one heading and words that begin with A are under another.
@property WordE119 headingSeparator;  // Returns or sets the text between alphabetic groups, entries that start with the same letter in the index.  
@property WordE105 indexFilter;  // Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index. 
@property WordE214 indexType;  // Returns or sets the index type.
@property NSInteger numberOfColumns;  // Sets or returns the number of columns for each page of an index.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in an index. 
@property WordE106 sortBy;  // Returns or sets the sorting criteria for the specified index.
@property WordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in an index. 
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the index


@end

// Represents a custom key assignment in the current context.
@interface WordKeyBinding : WordBaseObject

@property (copy, readonly) SBObject *bindingContext;  // Returns an object that represents the storage location of the specified key binding. This property can return a document or template object.
@property (copy, readonly) NSString *bindingKeyString;  // Returns the key combination string for the specified keys.
@property (copy, readonly) NSString *command;  // Returns the command assigned to the specified key combination.
@property (copy, readonly) NSString *commandParameter;  // Returns the command parameter assigned to the specified shortcut key.
@property (readonly) WordE239 keyCategory;  // Returns the type of item assigned to the specified key binding.
@property (readonly) NSInteger keyCode;  // Returns a unique number for the first key in the specified key binding.  You create this number by using the build key code method.
@property (readonly) NSInteger key_code_2;  // Returns a unique number for the second key in the specified key binding. You create this number by using the build key code method.
- (BOOL) protected;  // Returns true if you cannot change the specified key binding in the customize keyboard dialog box. 

- (void) disable;  // Removes the specified key combination if it's currently assigned to a command. After you use this method, the key combination has no effect.
- (void) executeKeyBinding;  // Runs the command associated with the specified key combination.
- (void) rebindKeyCategory:(WordE239)keyCategory command:(NSString *)command commandParameter:(NSString *)commandParameter;  // Changes the command assigned to the specified key binding.

@end

// Represents the elements of a letter created by the letter wizard.
@interface WordLetterContent : WordBaseObject

@property (copy) NSString *attentionLine;  // Returns or sets the attention line text for a letter created by the letter wizard.
@property (copy) NSString *ccList;  // Returns or sets the carbon copy recipients for a letter created by the letter wizard.
@property (copy) NSString *closing;  // Returns or sets the closing text for a letter created by the letter wizard, for example, Sincerely yours.
@property (copy) NSString *dateFormat;  // Returns or sets the date for a letter created by the letter wizard. 
@property NSInteger enclosureCount;  // Returns or sets the number of enclosures for a letter created by the letter wizard. 
@property BOOL includeHeaderFooter;  // Returns or sets if the header and footer from the page design template are included in a letter created by the letter wizard.
@property WordE245 letterStyle;  // Returns or sets the layout of a letter created by the letter wizard.
@property BOOL letterhead;  // Returns or sets if space is reserved for a preprinted letterhead in a letter created by the letter wizard.
@property WordE246 letterheadLocation;  // Returns or sets the location of the preprinted letterhead in a letter created by the letter wizard.
@property double letterheadSize;  // Returns or sets the amount of space in points to be reserved for a preprinted letterhead in a letter created by the letter wizard.
@property (copy) NSString *mailingInstructions;  // Returns or sets the mailing instruction text for a letter created by the letter wizard, for example, Certified Mail.
@property (copy) NSString *pageDesign;  // Returns or sets the name of the template attached to the document created by the letter wizard.
@property (copy) NSString *recipientAddress;  // Returns or sets the return address for a letter created with the letter wizard.
@property (copy) NSString *recipientName;  // Returns or sets the name of the person who'll be receiving the letter created by the letter wizard.
@property (copy) NSString *recipientReference;  // Returns or sets the reference line, for example, In reply to: for a letter created by the letter wizard.
@property (copy) NSString *returnAddress;  // Returns or sets the return address for a letter created with the letter wizard. 
@property (copy) NSString *salutation;  // Returns or sets the salutation text for a letter created by the letter wizard.
@property WordE247 salutationType;  // Returns or sets the type of salutation for a letter created by the letter wizard.  
@property (copy, readonly) NSString *senderCity;
@property (copy) NSString *senderCompany;  // Returns or sets the company name of the person creating a letter with the letter wizard.
@property (copy) NSString *senderInitials;  // Returns or sets the initials of the person creating a letter with the letter wizard. 
@property (copy) NSString *senderJobTitle;  // Returns or sets the job title of the person creating a letter with the letter wizard.
@property (copy) NSString *senderName;  // Returns or sets the name of the person creating a letter with the letter wizard. 
@property (copy) NSString *subject;  // Returns or sets the subject text of a letter created by the letter wizard.


@end

// Represents line numbers in the left margin or to the left of each newspaper-style column.
@interface WordLineNumbering : WordBaseObject

@property BOOL activeLine;  // Returns or sets if line numbering is active for the specified document, section, or sections.
@property NSInteger countBy;  // Returns or sets the numeric increment for line numbers. For example, if the count by property is set to 5, every fifth line will display the line number. Line numbers are only displayed in print layout view and print preview.
@property double distanceFromText;  // Returns or sets the distance in points between the right edge of line numbers and the left edge of the document text.
@property WordE173 restartMode;  // Returns or sets the way line numbering runs— that is, whether it starts over at the beginning of a new page or section or runs continuously.
@property NSInteger startingNumber;  // Returns or sets the starting line number.


@end

// Represents the linking characteristics for a picture.
@interface WordLinkFormat : WordBaseObject

@property BOOL autoUpdate;  // Returns or sets if the specified link is updated automatically when the container file is opened or when the source file is changed.
@property (readonly) WordE200 linkType;  // Returns the link type.
@property BOOL locked;  // Returns or sets if inline shape object is locked to prevent automatic updating.
@property BOOL savePictureWithDocument;  // Returns or sets if the specified picture is saved with the document.
@property (copy) NSString *sourceFullName;  // Returns or sets the path and name of the source file for the specified picture.
@property (copy, readonly) NSString *sourceName;  // Returns the name of the source file for the specified picture.
@property (copy, readonly) NSString *sourcePath;  // Returns the path of the source file for the specified picture.

- (void) breakLink;  // Breaks the link between the source file and the specified picture.

@end

// Represents an item in a drop-down form field.
@interface WordListEntry : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of this list entry object.


@end

// Represents the list formatting attributes that can be applied to the paragraphs in a range.
@interface WordListFormat : WordBaseObject

@property (copy, readonly) WordWordList *WordList;  // Returns a Word list object that represents the first formatted list contained in the list format object.
@property NSInteger listLevelNumber;  // Returns or sets the list level for the first paragraph in the list format object.
@property (copy) WordInlineShape *listPictureBullet;
@property (copy, readonly) NSString *listString;  // Returns a string that represents the appearance of the list value of the first paragraph in the range for the list format object. For example, the second paragraph in an alphabetic list would return B.
@property (copy, readonly) WordListTemplate *listTemplate;  // Returns a list template object that represents the list formatting for the list format object.
@property (readonly) WordE159 listType;  // Returns the type of lists that are contained in the range for the list format object.
@property (readonly) NSInteger listValue;  // Returns the numeric value of the first paragraph in the range for the specified list format object. For example, the list value property applied to the second paragraph in an alphabetic list would return 2.
@property (readonly) BOOL singleList;  // Returns if the specified list format object contains only one list.
@property (readonly) BOOL singleListTemplate;  // True if the entire list format object uses the same list template

- (void) applyBulletDefaultDefaultListBehavior:(WordE289)defaultListBehavior;  // Adds bullets and formatting to the paragraphs in the range for the list format object. If the paragraphs are already formatted with bullets, this method removes the bullets and formatting.
- (void) applyListFormatTemplateListTemplate:(WordListTemplate *)listTemplate continuePreviousList:(BOOL)continuePreviousList applyTo:(WordE137)applyTo defaultListBehavior:(WordE289)defaultListBehavior;  // Applies a set of list-formatting characteristics to the list format object
- (void) applyNumberDefaultDefaultListBehavior:(WordE289)defaultListBehavior;  // Adds the default numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.
- (void) applyOutlineNumberDefaultDefaultListBehavior:(WordE289)defaultListBehavior;  // Adds the default outline-numbering scheme to the paragraphs in the range for the list format object. If the paragraphs are already formatted as an outline-numbered list, this method removes the numbers and formatting.
- (void) listIndent;  // Increases the list level of the paragraphs in the range for the list format object, in increments of one level.
- (void) listOutdent;  // Decreases the list level of the paragraphs in the range for the list format object, in increments of one level.

@end

// Represents a single gallery of list formats.
@interface WordListGallery : WordBaseObject

- (SBElementArray *) listTemplates;

- (BOOL) modifiedIndex:(NSInteger)index;  // if the specified list template is not the built-in list template for that position in the list gallery. Index goes from 1 to 7
- (void) resetListGalleryIndex:(NSInteger)index;  // Resets the list template specified by index for the specified list gallery to the built-in list template format.

@end

// Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list.
@interface WordListLevel : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this list level object.
@property (copy) NSString *linkedStyle;  // Returns or sets the name of the style that's linked to the specified list level object.
@property WordE143 listLevelAlignment;  // Returns or sets a constant that represents the alignment for the list level of the list template.
@property (copy) NSString *numberFormat;  // Returns or sets the number format for the specified list level.
@property double numberPosition;  // Returns or sets the position in points of the number or bullet for the specified list level object.
@property WordE151 numberStyle;  // Returns or sets the number style for the list level object.
@property (copy, readonly) WordInlineShape *pictureBullet;
@property NSInteger resetOnHigher;  // Returns or sets the list level that must appear before the specified list level restarts numbering at 1.
@property NSInteger startAt;  // Returns or sets the starting number for the specified list level object.
@property double tabPosition;  // Returns or sets the tab position for the specified list level object.
@property double textPosition;  // Returns or sets the position in points for the second line of wrapping text for the specified list level object.
@property WordE149 trailingCharacter;  // Returns or sets the character inserted after the number for the specified list level.

- (WordInlineShape *) applyPictureBulletPath:(NSString *)path;

@end

// Represents a single list template that includes all the formatting that defines a list.
@interface WordListTemplate : WordBaseObject

- (SBElementArray *) listLevels;

@property (copy) NSString *name;  // Returns or sets the name of this list template object.
@property BOOL outlineNumbered;  // Returns or sets if the specified list template object is outline numbered.

- (WordListTemplate *) convertLevel:(NSInteger)level;  // Converts a multiple-level list to a single-level list, or vice versa.

@end

// Represents a mailing label.
@interface WordMailingLabel : WordBaseObject

- (SBElementArray *) customLabels;

@property (copy) NSString *defaultLabelName;  // Returns or sets the name for the default mailing label.
@property WordE207 defaultLaserTray;  // Returns or sets the default paper tray that contains sheets of mailing labels
@property BOOL defaultPrintBarCode;  // Returns or sets if a POSTNET bar code is added to envelopes or mailing labels by default. For U.S. mail only. This property must be set to true before the default print FIMA property is set

- (WordDocument *) createNewMailingLabelDocumentName:(NSString *)name address:(NSString *)address autoText:(NSString *)autoText extractAddress:(BOOL)extractAddress laserTray:(WordE207)laserTray singleLabel:(BOOL)singleLabel row:(NSInteger)row column:(NSInteger)column;  // Creates a new label document using either the default label options or ones that you specify. Returns a document object that represents the new document.
- (void) printOutMailingLabelName:(NSString *)name address:(NSString *)address extractAddress:(BOOL)extractAddress laserTray:(WordE207)laserTray singleLabel:(BOOL)singleLabel row:(NSInteger)row column:(NSInteger)column;  // Prints a label or a page of labels with the same address.

@end

// Represents an equation that has an accent mark above the base.
@interface WordMathAccent : WordBaseObject

- (NSInteger) character;  // Gets or sets an integer that represents the accent character for the accent object.
- (void) setChar: (NSInteger) character;
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.


@end

// Represents an individual entry in the Math AutoCorrect engine.
@interface WordMathAutocorrectEntry : WordBaseObject

@property (copy) NSString *autocorrectValue;  // Gets or sets a string that represents the contents of an equation auto correct entry.
@property (readonly) NSInteger entry_index;  // Gets an integer that represents the position of an item in the collection.
@property (copy) NSString *name;  // Gets or sets a string that represents the name of an equation auto correct entry.


@end

// Represents the Math AutoCorrect feature in Microsoft Word.
@interface WordMathAutocorrect : WordBaseObject

- (SBElementArray *) mathAutocorrectEntries;
- (SBElementArray *) mathRecognizedFunctions;

@property BOOL replaceText;  // Gets or sets whether Microsoft Word automatically replaces strings in equations with the corresponding math AutoCorrect definitions.
@property BOOL useOutsideEquations;  // Gets or sets whether Microsoft Word uses math autocorrect rules outside equations in a document.

- (void) addMathAcEntryName:(NSString *)name value:(NSString *)value;  // Creates an equation auto correct entry.
- (void) addMathRecognizedFunctionName:(NSString *)name;  // Creates a new recognized function.

@end

// Represents the mathematical overbar for an object in an equation.
@interface WordMathBar : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isTopBar;  // Gets or sets the position of a bar in a bar object. True specifies a mathematical overbar. False specifies a mathematical underbar


@end

// Represents an invisible box around an equation or part of an equation to which you can assign properties that affect the layout or mathematical formatting of the entire box.
@interface WordMathBorderBox : WordBaseObject

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
@interface WordMathBox : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isDifferential;  // Gets or sets whether the box acts as the mathematical differential, in which case the box receives the appropriate horizontal spacing for a differential.
@property BOOL isOperatorEmulator;  // Gets or sets if the box and its contents behave as a single operator and inherits the properties of an operator.
@property BOOL nonBreaking;  // Gets or sets whether breaks are allowed inside the box object.


@end

// Represents individual line breaks in an equation.
@interface WordMathBreak : WordBaseObject

@property NSInteger alignAt;  // Gets or sets an integer that represents the operator in one line, to which to align consecutive lines in an equation.
@property (copy, readonly) WordTextRange *textRange;  // Returns a text range object that represents the portion of a document that is contained in the specified object.


@end

// Represents a delimiter object, consisting of opening and closing delimiters, e.g. parentheses, braces, brackets, or vertical bars, and one or more elements contained inside the delimiters.
@interface WordMathDelimiter : WordBaseObject

- (SBElementArray *) mathObjects;

@property NSInteger beginningCharacter;  // Gets or sets an Integer that represents the beginning delimiter character in a delimiter object.
@property WordE325 delimiterShape;  // Gets or sets the appearance of delimiters, e.g. parentheses, braces, and brackets, in relationship to the content that they surround.
@property NSInteger endingCharacter;  // Gets or sets an Integer that represents the ending delimiter character in a delimiter object.
@property BOOL hideLeftDelimiter;  // Gets or sets whether to hide the opening delimiter in a delimiter object.
@property BOOL hideRightDelimiter;  // Gets or sets whether to hide the closing delimiter in a delimiter object.
@property NSInteger separatorCharacter;  // Gets or sets an Integer that represents the separator character in a delimiter object when the delimiter object contains two or more arguments.
@property BOOL shouldGrow;  // Gets or sets whether delimiter characters grow to the full height of the arguments that they contain.


@end

// Represents a mathematical equation array object, consisting of one or more equations that can be vertically justified as a unit respect to surrounding text on the line.
@interface WordMathEquationArray : WordBaseObject

- (SBElementArray *) mathObjects;

@property BOOL distributeEqually;  // Gets or sets if the equations in an equation array are distributed equally within the margins of its container, such as a column, cell, or page width.
@property NSInteger rowSpacing;  // Gets or sets an integer that represents the spacing between the rows in an equation array.
@property WordE323 rowSpacingRule;  // Gets or sets the spacing rule that defines spacing in an equation array.
@property BOOL useMaxWidth;  // Gets or sets whether the equations in an equation array are spaced to the maximum width of the equation array.
@property WordE321 verticalAlignment;  // Gets or sets the type of vertical alignment for an equation array with respect to the text that surrounds the array.


@end

// Represents a fraction, consisting of a numerator and denominator separated by a fraction bar. The fraction bar can be horizontal or diagonal, depending on the fraction properties.
@interface WordMathFraction : WordBaseObject

@property (copy, readonly) WordMathObject *denominator;  // Returns a math object object that represents the denominator for an equation that contains a fraction.
@property WordE322 fractionType;  // Gets or sets the layout of a fraction, whether it is stacked, skewed, linear, or without a fraction bar.
@property (copy, readonly) WordMathObject *numerator;  // Returns a math object object that represents the numerator for the fraction.


@end

// Represents a math func object that represents a type of mathematical function that consists of a function name, such as sin or cos, and an argument.
@interface WordMathFunc : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *funcName;  // Returns an OMath object that represents the name of a mathematical function, such as sin or cos.


@end

// Represents a mathematical function or structure such as fractions, integrals, sums, and radicals.
@interface WordMathFunction : WordBaseObject

- (SBElementArray *) mathObjects;

@property (copy, readonly) WordMathAccent *accent;  // Returns a math accent object that represents a base character with a combining accent mark.
@property (copy, readonly) WordMathBorderBox *borderBox;  // Returns a math border box object that represents a border drawn around an equation or part of an equation.
@property (copy, readonly) WordMathBox *box;  // Returns a math box object to which you can apply properties.
@property (copy, readonly) WordMathDelimiter *delimiter;  // Returns an math delimiter object that represents the delimiter function
@property (copy, readonly) WordMathEquationArray *equationArray;  // Returns a math equation array object that represents an equation array function.
@property (copy, readonly) WordMathFraction *fraction;  // Returns a math fraction object that represents a fraction.
@property (copy, readonly) WordMathFunc *func;  // Returns a math func object that represents a type of mathematical function.
@property (readonly) WordE319 functionType;  // Gets the type of the function.
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
@interface WordMathGroupChar : WordBaseObject

@property BOOL alignTop;  // Returns or sets whether the grouping character is aligned vertically with the surrounding text or whether the base text that is either above or below the grouping character is aligned vertically with the surrounding text.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL isOnTop;  // Gets or sets whether the grouping character is placed above the base text of the group character object. False displays the group character under the base text.
@property NSInteger theCharacter;  // Gets or sets an integer that represents the character placed above or below text in a group character object.


@end

// Represents an equation that contains a superscript or subscript to the left of the base.
@interface WordMathLeftScripts : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *subscript;  // Returns a math object object that represents the subscript for a pre-sub-superscript object.
@property (copy, readonly) WordMathObject *superscript;  // Returns a math object object that represents the superscript for a pre-sub-superscript object.

- (WordMathFunction *) convertLeftScriptsToSubAndSuperScripts;  // Converts an equation with a superscript or subscript to the left of the base of the equation to an equation with a base of a superscript or subscript.

@end

// Represents the lower limit mathematical construct, consisting of text on the baseline and reduced-size text immediately below it.
@interface WordMathLowerLimit : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *limit;  // Returns a math object object that represents the limit of the lower limit object. 

- (WordMathFunction *) lowerLimitToUpperLimit;

@end

// Represents a matrix column.
@interface WordMathMatrixColumn : WordBaseObject

- (SBElementArray *) mathObjects;

@property NSInteger columnIndex;  // Gets an integer that represents the ordinal position of a matrix column within the collection of matrix columns.
@property WordE320 horizontalAlignment;  // Gets or sets the horizontal alignment for arguments in a matrix column.


@end

// Represents a matrix row.
@interface WordMathMatrixRow : WordBaseObject

- (SBElementArray *) mathObjects;

@property NSInteger rowIndex;  // Gets an integer that represents the ordinal position of a matrix row within the collection of matrix rows.


@end

// Represents an equation matrix.
@interface WordMathMatrix : WordBaseObject

- (SBElementArray *) mathMatrixRows;
- (SBElementArray *) mathMatrixColumns;

@property NSInteger columnGap;  // Gets or sets an integer that represents the spacing between columns in a matrix.
@property WordE323 columnGapRule;  // Gets or sets the spacing rule for the space that appears between columns in a matrix.
@property NSInteger columnSpacing;  // Gets or sets an integer that represents the spacing for columns in a matrix.
@property BOOL hidePlaceholders;  // Gets or sets whether placeholders in a matrix are hidden from display.
@property NSInteger rowSpacing;  // Gets or sets an integer that represents the spacing for rows in a matrix. 
@property WordE323 rowSpacingRule;  // Gets or sets the spacing rule for rows in a matrix.
@property WordE321 verticalAlignment;  // Gets or sets the vertical alignment for a matrix.

- (WordMathMatrixColumn *) addMatrixColumnBeforeColumn:(WordMathMatrixColumn *)beforeColumn;  // Creates a matrix column and adds it to a matrix and returns a math matrix column object.
- (WordMathMatrixRow *) addMatrixRowBeforeRow:(WordMathMatrixRow *)beforeRow;  // Creates a matrix row and adds it to a matrix and returns a math matrix row object.

@end

// Represents the mathematical n-ary object, consisting of an n-ary object, a base/operand, and optional upper limits and lower limits.
@interface WordMathNary : WordBaseObject

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
@interface WordMathObject : WordBaseObject

- (SBElementArray *) mathFunctions;
- (SBElementArray *) mathBreaks;

@property NSInteger alignPoint;  // Gets or sets an integer that represents the character position of the alignment point in the equation.
@property (readonly) NSInteger argumentIndex;  // Gets an integer that specifies the argument index of this component relative to the containing math object.
@property (copy, readonly) WordMathFunction *containingFunction;  // Gets the math function that represents the parent, or containing, function.
@property WordE324 displayType;  // Sets or returns the display type of the equation.
@property WordE326 justification;  // Gets or sets the justification for the equation.
@property (readonly) NSInteger nestingLevel;  // Returns an integer that represents the nesting level for the math object.
@property (copy, readonly) WordMathObject *parentArgument;  // Gets a math object that represents the parent, or containing, argument.
@property (copy, readonly) WordMathMatrixColumn *parentColumn;  // Gets a math matrix column object that represents the parent column in a matrix.
@property (copy, readonly) WordMathObject *parentObject;  // Gets the math object that represents the parent object.
@property (copy, readonly) WordMathMatrixRow *parentRow;  // Gets a math matrix row object that represents the parent row in a matrix.
@property NSInteger scriptSize;  // Gets or sets an integer that represents the script size of an argument, for example, text, script, or script-script.
@property (copy, readonly) WordTextRange *textRange;  // Gets a text range object that represents the portion of a document that contains the equation.

- (void) addMathFunctionInLocation:(WordTextRange *)inLocation mathFunctionType:(WordE319)mathFunctionType numberOfArguments:(NSInteger)numberOfArguments numberOfColumns:(NSInteger)numberOfColumns;  // Inserts a new structure, such as a fraction, into an equation at the specified position.
- (void) buildUp;  // Converts an equation to professional/built up format.
- (void) convertToLiteralText;  // Converts an equation to literal text.
- (void) convertToMathText;  // Converts an equation to math text.
- (void) convertToNormalText;  // Converts an equation to normal text.
- (WordMathBreak *) insertMathBreakAtRange:(WordTextRange *)atRange;  // Inserts a break into an equation and returns a math break object that represents the break.
- (void) linearize;  // Converts an equation to linear/built down format.

@end

// Represents a phantom object, which has two primary uses: 1. adding the spacing of the phantom base without displaying that base or 2. suppressing part of the glyph from spacing considerations.
@interface WordMathPhantom : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL showsInvisibles;  // Gets or sets whether the contents of a phantom object are visible.
@property BOOL smash;  // Gets or sets if the contents of the phantom are visible but that the height is not taken into account in the spacing of the layout.
@property BOOL transparent;  // Gets or sets whether a phantom object is transparent.
@property BOOL zeroAscent;  // Gets or sets whether the ascent of the phantom contents is ignored in the spacing of the layout.
@property BOOL zeroDescent;  // Gets or sets whether the descent of the phantom contents is ignored in the spacing of the layout.
@property BOOL zeroWidth;  // Gets or sets whether the width of a phantom object is ignored in the spacing of the layout.


@end

// Represents the mathematical radical object, consisting of a radical, a base, and an optional degree.
@interface WordMathRadical : WordBaseObject

@property (copy, readonly) WordMathObject *degree;  // Returns a math object object that represents the degree for a radical.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property BOOL hideDegree;  // Gets or sets whether to hide the degree for the radical.


@end

@interface WordMathRecognizedFunction : WordBaseObject

@property (readonly) NSInteger entry_index;  // Gets an integer that represents the position of an item in the collection.
@property (copy, readonly) NSString *name;  // Gets a string that represents the name of an equation recognized function.


@end

// Represents an equation with a base that contains a superscript or subscript.
@interface WordMathSubAndSuperScript : WordBaseObject

@property BOOL alignScripts;  // Gets or sets whether to horizontally align subscripts and superscripts in the sub-superscript object.
@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *subscript;  // Returns a math object object that represents the subscript for a subscript-superscript object.
@property (copy, readonly) WordMathObject *superscript;  // Returns a math object object that represents the superscript for a subscript-superscript object.

- (WordMathFunction *) convertSubAndSuperScriptsToLeftScripts;  // Converts an equation with a base superscript or subscript to an equation with a superscript or subscript to the left of the base.
- (WordMathFunction *) removeSubscript;  // Removes the subscript for an equation and returns a math function object that represents the updated equation without the subscript.
- (WordMathFunction *) removeSuperscript;  // Removes the superscript for an equation and returns a math function object that represents the updated equation without the superscript.

@end

// Represents an equation with a base that contains a subscript.
@interface WordMathSubscript : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *subscript;  // Returns a math object that represents the subscript for a subscript object.


@end

// Represents an equation with a base that contains a superscript.
@interface WordMathSuperscript : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *superscript;  // Returns a math object object that represents the superscript for a superscript object.


@end

// Represents the upper limit mathematical construct, consisting of text on the baseline and reduced-size text immediately above it.
@interface WordMathUpperLimit : WordBaseObject

@property (copy, readonly) WordMathObject *equationBase;  // Returns a math object object that represents the base of the equation object.
@property (copy, readonly) WordMathObject *limit;  // Returns a math object object that represents the limit of the upper limit object.

- (WordMathFunction *) upperLimitToLowerLimit;

@end

// Represents options associated with page number objects
@interface WordPageNumberOptions : WordBaseObject

@property WordE120 chapterPageSeparator;  // Returns or sets the separator character used between the chapter number and the page number. 
@property NSInteger headingLevelForChapter;  // Returns or sets the heading level style that's applied to the chapter titles in the document. Can be a number from zero through 8, corresponding to heading levels 1 through 9.
@property BOOL includeChapterNumber;  // Returns or sets if a chapter number is included with page numbers.
@property WordE154 numberStyle;  // Returns or sets the number style for the page number objects.
@property BOOL restartNumberingAtSection;  // Returns or sets if page numbering starts at 1 again at the beginning of the specified section.
@property BOOL showFirstPageNumber;  // Returns or sets if the page number appears on the first page in the section.
@property NSInteger startingNumber;  // Returns or sets the starting page number.


@end

// Represents a page number in a header or footer.
@interface WordPageNumber : WordBaseObject

@property WordE121 alignment;  // Returns or sets a constant that represents the alignment for the page number.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.


@end

// Represents the page setup description.
@interface WordPageSetup : WordBaseObject

- (SBElementArray *) textColumns;

@property double bottomMargin;  // Returns or sets the distance in points between the bottom edge of the page and the bottom boundary of the body text.
@property double charsLine;  // Returns or sets the number of characters per line in the document grid.
@property BOOL differentFirstPageHeaderFooter;  // Returns or set if a different header or footer is used on the first page.
@property WordE207 firstPageTray;  // Returns or sets the paper tray to use for the first page of a document or section.
@property double footerDistance;  // Returns or sets the distance in points between the footer and the bottom of the page.
@property double gutter;  // Returns or sets the amount in points of extra margin space added to each page in a document or section for binding. 
@property WordE283 gutterPosition;  // Returns or sets on which side the gutter appears in a document.
@property double headerDistance;  // Returns or sets the distance in points between the header and the top of the page.
@property WordE278 layoutMode;  // Returns or sets the layout mode for the current document.
@property double leftMargin;  // Returns or sets the distance in points between the left edge of the page and the left boundary of the body text.
@property BOOL lineBetweenTextColumns;  // Returns or sets if vertical lines appear between all the columns.
@property (copy) WordLineNumbering *lineNumbering;  // Returns or sets the line numbering object that represents the line numbers for the specified page setup object.
@property double linesPage;  // Returns or sets the number of lines per page in the document grid.
@property BOOL mirrorMargins;  // Returns or sets if the inside and outside margins of facing pages are the same width.
@property BOOL oddAndEvenPagesHeaderFooter;  // Returns or sets if the specified page setup object has different headers and footers for odd-numbered and even-numbered pages.
@property WordE208 orientation;  // Returns or sets the orientation of the page.
@property WordE207 otherPagesTray;  // Returns or sets the paper tray to be used for all but the first page of a document or section.
@property double pageHeight;  // Returns or sets the height of the page in points.
@property double pageWidth;  // Returns or sets the width of the page in points.
@property WordE232 paperSize;  // Returns or sets the paper size.
@property double rightMargin;  // Returns or sets the distance in points between the right edge of the page and the right boundary of the body text.
@property WordE219 sectionStart;  // Returns or sets the type of section break for the specified object.
@property BOOL showGrid;  // Determines whether to show the grid.
@property double spacingBetweenTextColumns;  // Returns or sets the spacing in points between columns.
@property BOOL suppressEndnotes;  // Returns or sets if endnotes are printed at the end of the next section that doesn't suppress endnotes. Suppressed endnotes are printed before the endnotes in that section.
@property BOOL textColumnsEvenlySpaced;  // Returns or sets if text columns are evenly spaced.
@property double topMargin;  // Returns or sets the distance in points between the top edge of the page and the top boundary of the body text.
@property WordE146 verticalAlignment;  // Returns or sets the vertical alignment of text on each page in a document or section.
@property double widthOfTextColumns;  // Returns or sets the width of all text columns

- (void) setAsPageSetupTemplateDefault;  // Sets the specified page setup formatting as the default for the active document and all new documents based on the active template.
- (void) setNumberOfTextColumnsNumberOfColumns:(NSInteger)numberOfColumns;  // Arranges text into the specified number of text columns. 
- (void) togglePortrait;  // Switches between portrait and landscape page orientations for a document or section.

@end

// Represents a window pane.
@interface WordPane : WordBaseObject

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

- (WordZoom *) getZoomZoomType:(WordE202)zoomType;  // Returns a zoom object of the specified type for this pane.

@end

@interface WordRangeEndnoteOptions : WordBaseObject

@property WordE175 endnoteLocation;  // Returns or sets the position of endnotes in a range or selection.
@property WordE152 endnoteNumberStyle;  // Returns or sets the number style for endnotes in a range or selection.
@property WordE173 endnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks.
@property NSInteger endnoteStartingNumber;  // Returns or sets the starting endnote number in a range or selection.


@end

@interface WordRangeFootnoteOptions : WordBaseObject

@property WordE174 footnoteLocation;  // Returns or sets the position of footnotes in a range or selection.
@property WordE152 footnoteNumberStyle;  // Returns or sets the number style for footnotes in a range or selection.
@property WordE173 footnoteNumberingRule;  // Returns or sets the way footnotes or endnotes are numbered after page breaks or section breaks. 
@property NSInteger footnoteStartingNumber;  // Returns or sets the starting number for footnotes in a range or selection.


@end

// Represents a recently used file.
@interface WordRecentFile : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the recently used file.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the recent file.
@property BOOL readOnly;  // Returns or sets if changes to the document cannot be saved to the original document. 

- (WordDocument *) openRecentFile;  // Opens the recent file and returns a document object.

@end

// Represents the replace criteria for a find-and-replace operation.
@interface WordReplacement : WordBaseObject

@property (copy) NSString *content;  // Returns or sets the text to replace in the specified range or selection.
@property (copy, readonly) WordFont *fontObject;  // The font object associated with this replacement object.
@property (copy, readonly) WordFrame *frame;  // The frame object associated with this replacement object.
@property NSInteger highlight;  // Returns or sets if highlight formatting is applied to the replacement text.
@property WordE182 languageID;  // Returns or sets the language for the replacement object
@property WordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or set the paragraph format object associated with this replacement object.
@property WordE184 style;  // Returns or sets the Word style associated with the replacement object.


@end

// Property of View: a person who has made tracked changes in the viewed document.
@interface WordReviewer : WordBaseObject

@property BOOL visible;


@end

// Represents a change marked with a revision mark.
@interface WordRevision : WordBaseObject

@property (copy, readonly) NSString *author;  // Returns the name of the user who made the specified tracked change. 
@property (copy, readonly) id cells;
@property (copy, readonly) NSDate *dateValue;  // The date and time that the tracked change was made.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *formatDescription;
@property (copy, readonly) WordTextRange *movedRange;
@property (readonly) WordE216 revisionType;  // Returns the revision type.
@property (copy, readonly) id style;
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the revision


@end

// Represents the current selection in a window or pane. A selection represents either a selected or highlighted area in the document, or it represents the insertion point if nothing in the document is selected.
@interface WordSelectionObject : WordBaseObject

- (SBElementArray *) tables;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) characters;
- (SBElementArray *) footnotes;
- (SBElementArray *) endnotes;
- (SBElementArray *) WordComments;
- (SBElementArray *) cells;
- (SBElementArray *) sections;
- (SBElementArray *) paragraphs;
- (SBElementArray *) fields;
- (SBElementArray *) formFields;
- (SBElementArray *) frames;
- (SBElementArray *) bookmarks;
- (SBElementArray *) hyperlinkObjects;
- (SBElementArray *) columns;
- (SBElementArray *) rows;
- (SBElementArray *) inlineShapes;
- (SBElementArray *) shapes;
- (SBElementArray *) mathObjects;

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
@property WordE182 languageID;  // Returns or sets the language for the selection object
@property WordE182 languageIDEastAsian;  // Returns or sets an east asian language for the selection.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores.
@property WordE270 orientation;  // Returns or sets the orientation of text in the selection when the text direction feature is enabled.
@property (copy) WordPageSetup *pageSetup;  // Returns or set the page setup object associated with the selection.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or set the paragraph object associated with the selection.
@property (readonly) NSInteger previousBookmarkId;  // Returns the number of the last bookmark that starts before or at the same place as the selection, It returns zero if there's no corresponding bookmark.
@property (copy, readonly) WordRangeEndnoteOptions *rangeEndnoteOptions;
@property (copy, readonly) WordRangeFootnoteOptions *rangeFootnoteOptions;
@property (copy, readonly) WordRowOptions *rowOptions;
@property NSInteger selectionEnd;  // Returns or sets the ending character position of the selection.
@property WordE261 selectionFlags;  // Returns or sets properties of the selection.
@property (readonly) BOOL selectionIsActive;  // Returns if the selection is active.
@property NSInteger selectionStart;  // Returns or sets the starting character position of the selection.
@property (readonly) WordE209 selectionType;  // Returns the selection type.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the selection
@property (copy) NSString *showWordCommentsBy;  // Returns or sets the name of the reviewer whose comments are shown in the comments pane. You can choose to show comments either by a single reviewer or by all reviewers. To view the comments by all reviewers, set this property to 'All Reviewers'.
@property BOOL showHiddenBookmarks;  // Returns or sets if hidden bookmarks are included in the elements of the selection
@property BOOL startIsActive;  // Returns or sets if the beginning of the selection is active. If the selection is not collapsed to an insertion point, either the beginning or the end of the selection is active.
@property (readonly) NSInteger storyLength;  // Returns the number of characters in the story that contains the selection
@property (readonly) WordE160 storyType;  // Returns the story type for the selection.
@property WordE184 style;  // Returns or sets the Word style associated with the selection object.
@property WordE182 supplementalLanguageID;  // Returns or sets the language for the selection
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the selection

- (double) calculateSelection;  // Calculates a mathematical expression within a selection.
- (void) copyFormat;  // Copies the character formatting of the first character in the selected text. If a paragraph mark is selected, Word copies paragraph formatting in addition to character formatting.
- (void) createTextbox;  // Adds a default-size text box around the selection. If the selection is an insertion point, this method changes the pointer to a cross-hair pointer so that the user can draw a text box.
- (WordTextRange *) endKeyMove:(WordE295)move extend:(WordE249)extend;  // Moves or extends the selection to the end of the specified unit. This method returns the new range object or missing value if an error occurred.
- (void) escapeKey;  // Cancels a mode such as extend or column select.
- (id) expandBy:(WordE129)by;  // Expands the selection. Returns the number of characters added to the range or selection.
- (void) extendCharacter:(NSString *)character;  // Extends the selection to the next larger unit of text. The progression of selected units of text is as follows: word, sentence, paragraph, section, entire document. The selection is extended by moving the active end of the selection.
- (WordField *) getNextField;  // Returns the next field object.
- (WordField *) getPreviousField;  // Returns the previous field object.
- (NSString *) getSelectionInformationInformationType:(WordE266)informationType;  // Returns information about the selection. 
- (NSInteger) homeKeyMove:(WordE295)move extend:(WordE249)extend;  // Moves or extends the selection to the beginning of the specified unit. This method returns the new range object or missing value if an error occurred.
- (void) insertCellsShiftCells:(WordE135)shiftCells;  // Adds cells to an existing table. The number of cells inserted is equal to the number of cells in the selection.
- (void) insertColumnsPosition:(WordEiRL)position;  // Inserts columns to the left of the column that contains the selection. If the selection isn't in a table, an error occurs. The number of columns inserted is equal to the number of columns selected.
- (void) insertFormulaFormula:(NSString *)formula numberFormat:(NSString *)numberFormat;  // Inserts an = Formula field that contains a formula at the selection. The formula replaces the selection, if the selection isn't collapsed.
- (void) insertRowsPosition:(WordEiAb)position numberOfRows:(NSInteger)numberOfRows;  // Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.
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
@interface WordSubdocument : WordBaseObject

@property (readonly) BOOL hasFile;  // Returns true if the specified subdocument has been saved to a file.
@property (readonly) NSInteger level;  // Returns the heading level used to create the subdocument.
@property BOOL locked;  // Returns or sets if a subdocument in a master document is locked.
@property (copy, readonly) NSString *name;  // The name of the subdocument.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified subdocument.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the subdocument.

- (WordDocument *) openSubdocument;  // Opens the specified subdocument. Returns a document object representing the opened object.
- (void) splitSubdocumentTextRange:(WordTextRange *)textRange;  // Divides an existing subdocument into two subdocuments at the same level in master document view or outline view. The division is at the beginning of the specified range.

@end

// Contains information about the computer system.
@interface WordSystemObject : WordBaseObject

@property (readonly) WordE118 country;  // Returns the country or region designation of the system
@property WordE139 cursor;  // Returns or sets the state of the pointer.
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
@interface WordTabStop : WordBaseObject

@property WordE145 alignment;  // Returns or sets a constant that represents the alignment for the specified tab stop.
@property (readonly) BOOL customTab;  // Returns if the specified tab stop is a custom tab stop.
@property (copy, readonly) WordTabStop *nextTabStop;  // Returns the next tab stop object
@property (copy, readonly) WordTabStop *previousTabStop;  // Returns the previous tab stop object
@property WordE170 tabLeader;  // Returns or sets the leader for the specified tab stop object
@property double tabStopPosition;  // Returns or sets the position of a tab stop relative to the left margin.


@end

// Represents a single table of authorities in a document
@interface WordTableOfAuthorities : WordBaseObject

@property NSInteger category;  // Returns or sets the category of entries to be included in a table of authorities.
@property (copy) NSString *entrySeparator;  // Returns or sets the characters up to five that separate a table of authorities entry and its page number. The default is a tab character with a dotted leader.
@property BOOL includeCategoryHeader;  // Returns or sets if the category name for a group of entries appears in the table of authorities.
@property (copy) NSString *includeSequenceName;  // Returns or sets the Sequence SEQ field identifier for a table of authorities.
@property BOOL keepEntryFormatting;  // Returns or sets if formatting from table of authorities entries is applied to the entries in the specified table of authorities.
@property (copy) NSString *pageNumberSeparator;  // Returns of sets the characters up to five that separate individual page references in a table of authorities. The default is a comma and a space.
@property (copy) NSString *pageRangeSeparator;  // Returns or sets the characters up to five that separate a range of pages in a table of authorities. The default is an en dash.
@property BOOL passim;  // Returns or sets if five or more page references to the same authority are replaced with Passim.
@property (copy) NSString *separator;  // Returns or sets the characters up to five between the sequence number and the page number. A hyphen is the default character.
@property WordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in a table of authorities.
@property (copy) NSString *tableOfAuthoritiesBookmark;  // Returns or sets the name of the bookmark from which to collect table of authorities entries.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the table of authorities object.


@end

// Represents a single table of contents in a document.
@interface WordTableOfContents : WordBaseObject

- (SBElementArray *) headingStyles;

@property BOOL includePageNumbers;  // Returns or sets if page numbers are included in the table of contents.
@property NSInteger lowerHeadingLevel;  // Returns or sets the ending heading level for a table of contents.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in a table of contents.
@property WordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in a table of contents.
@property (copy) NSString *tableId;  // Returns or sets a one-letter identifier that's used to build a table of contents or table of figures from TOC fields.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the table of contents object.
@property NSInteger upperHeadingLevel;  // Returns or sets the starting heading level for a table of contents.
@property BOOL useFields;  // Returns or sets if table of contents entry fields are used to create a table of contents.
@property BOOL useHeadingStyles;  // Returns or sets if built-in heading styles are used to create a table of contents.


@end

// Represents a single table of figures in a document.
@interface WordTableOfFigures : WordBaseObject

- (SBElementArray *) headingStyles;

@property (copy) NSString *caption;  // Returns or sets the label that identifies the items to be included in a table of figures.
@property BOOL includeLabel;  // Returns or sets if the caption label and caption number are included in a table of figures.
@property BOOL includePageNumbers;  // Returns or sets if page numbers are included in the table of figures.
@property NSInteger lowerHeadingLevel;  // Returns or sets the ending heading level for a table of figures.
@property BOOL rightAlignPageNumbers;  // Returns or sets if page numbers are aligned with the right margin in a table of figures.
@property WordE170 tabLeader;  // Returns or sets the character between entries and their page numbers in a table of figures.
@property (copy) NSString *tableId;  // Returns or sets a one-letter identifier that's used to build a table of figures from TOC fields.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the text for the table of figures object.
@property NSInteger upperHeadingLevel;  // Returns or sets the starting heading level for a table of figures.
@property BOOL useFields;  // Returns or sets if table of figures entry fields are used to create a table of figures.
@property BOOL useHeadingStyles;  // Returns or sets if built-in heading styles are used to create a table of figures.


@end

// Represents a document template.
@interface WordTemplate : WordBaseObject

- (SBElementArray *) autoTextEntries;
- (SBElementArray *) documentProperties;
- (SBElementArray *) customDocumentProperties;
- (SBElementArray *) listTemplates;
- (SBElementArray *) buildingBlocks;
- (SBElementArray *) buildingBlockTypes;

@property WordE108 eastAsianLineBreak;  // Returns or sets the line break control level for the specified template. This property is ignored if the “east asian line break control” property is set to false. Note that “east asian line break control” is a paragraph format property.
@property (copy, readonly) NSString *fullName;  // Specifies the name of the template including the drive or Web path.
@property WordE107 justificationMode;  // Returns or sets the character spacing adjustment for the specified template.
@property WordE182 languageID;  // Returns or sets the language for the template object
@property WordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (copy, readonly) NSString *name;  // Returns the name of the template.
@property (copy) NSString *noLineBreakAfter;  // Returns or sets the kinsoku characters after which Microsoft Word will not break a line.
@property (copy) NSString *noLineBreakBefore;  // Returns or sets the kinsoku characters before which Microsoft Word will not break a line.
@property BOOL noProofing;  // Returns or sets if the spelling and grammar checker should ignore documents based on this template.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the template file.
@property BOOL saved;  // Return or set if the template hasn't changed since it was last saved. False if Microsoft Word displays a prompt to save changes when the document is closed.
@property (readonly) WordE101 templateType;  // Returns the template type.

- (WordBuildingBlock *) addBuildingBlockToTemplateName:(NSString *)name type:(WordE329)type category:(NSString *)category fromLocation:(WordTextRange *)fromLocation objectDescription:(NSString *)objectDescription insertOptions:(WordE330)insertOptions;  // Creates a new building block entry in a template and returns a building block object that represents the new building block entry.
- (WordAutoTextEntry *) appendToSpikeRange:(WordTextRange *)range;  // Deletes the specified range and adds the contents of the range to the Spike which is a built-in autotext entry.
- (WordDocument *) openAsDocument;  // Opens the specified template as a document and returns a document object. Opening a template as a document allows the user to edit the contents of the template.

@end

// Represents a single text column.
@interface WordTextColumn : WordBaseObject

@property double spaceAfter;  // Returns or sets the amount of spacing in points after the text column.
@property double width;  // Returns or sets the width of the object.


@end

// Represents a single text form field.
@interface WordTextInput : WordBaseObject

@property (copy) NSString *defaultTextInput;  // Returns or sets the text that represents the default text box contents.
@property (copy, readonly) NSString *format;  // Returns the formatting string for this text input object.
@property (readonly) WordE188 textInputFieldType;  // Returns the type of text form field.
@property (readonly) BOOL valid;  // Returns if the text input object is valid.
@property NSInteger width;  // Returns or sets the width of the object.

- (void) editTypeFormFieldType:(WordE188)formFieldType defaultType:(NSString *)defaultType typeFormat:(NSString *)typeFormat enabled:(BOOL)enabled;  // Sets options for the specified text form field.

@end

// Represents options that control how text is retrieved from a text range object.
@interface WordTextRetrievalMode : WordBaseObject

@property BOOL includeFieldCodes;  // Returns or sets if the text retrieved from the specified range includes field codes.
@property BOOL includeHiddenText;  // Returns or sets if the text retrieved from the specified range includes hidden text.
@property WordE202 viewType;  // Returns or sets the view for the text retrieval mode object.


@end

// Represents a variable stored as part of a document. Document variables are used to preserve macro settings in between macro sessions.
@interface WordVariable : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of the variable
@property (copy) NSString *variableValue;  // Returns or sets the value of a variable as a string.


@end

// Contains the view attributes, show all, field shading, table gridlines, and so on, for a window or pane.
@interface WordView : WordBaseObject

- (SBElementArray *) reviewers;

@property BOOL browseToWindow;  // Returns or sets if lines wrap at the right edge of the window rather than at the right margin of the page.
@property BOOL conflictMode;  // Returns or sets the Conflict Mode.
@property BOOL dataMergeDataView;  // Returns or sets if data merge data is displayed instead of data merge fields in the specified window.
@property BOOL draft;  // Returns or sets if all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.
@property NSInteger enlargeFontsLessThan;  // Returns or sets the point size below which screen fonts are automatically scaled to the larger size.
@property WordE229 fieldShading;  // Returns or sets on-screen shading for form fields.
@property BOOL fullScreen;  // Returns or sets if the window is in full-screen view.
@property BOOL magnifier;  // Returns or sets if the pointer is displayed as a magnifying glass in print preview, indicating that the user can click to zoom in on a particular area of the page or zoom out to see an entire page or spread of pages.
@property WordE301 revisionsMode;  // Returns or sets the way Microsoft Word shows revisions.
@property WordE300 revisionsView;  // Returns or sets whether Microsoft Word shows revisions based on the final or the original document.
@property WordE203 seekView;  // Returns or sets the document element displayed in print layout view.
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
@property WordE204 splitSpecial;  // Returns or sets the active window pane.
@property BOOL tableGridlines;  // Returns or sets if table gridlines are displayed. 
@property WordE202 viewType;  // Returns or sets the view type.
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
@interface WordWebOptions : WordBaseObject

@property BOOL allowPng;  // Returns or sets if PNG, Portable Network Graphics, is allowed as an image format when you save a document as a Web page.
@property (copy, readonly) NSString *docKeywords;  // Returns the keywords associated with a document.
@property (copy, readonly) NSString *docTitle;  // Returns the title for a Web document.
@property WordMtEn encoding;  // Returns or sets the document encoding, code page or character set, to be used by the Web browser when you view the saved document
@property NSInteger pixelsPerInch;  // Returns or sets the density, pixels per inch, of graphics images and table cells on a Web page. The range of settings is usually from 19 to 480, and common settings for popular screen sizes are 72, 96, and 120.
@property BOOL roundTripHTML;  // Returns or sets if whether to save an HTML document with information that is specific to Microsoft Word. Setting this property to true allows you to preserve all Word settings in an HTML document.
@property WordMSsz screenSize;  // Returns or sets the ideal minimum screen size, width by height, in pixels, that you should use when viewing the saved document in a Web browser.
@property BOOL useLongFileNames;  // Returns or sets if long file names are used when you save the document as a Web page.

- (void) useDefaultFolderSuffix;  // Sets the folder suffix for the specified document to the default suffix for the language support you have selected or installed.

@end

// Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.
@interface WordWindow : WordBasicWindow

- (SBElementArray *) panes;

@property WordE103 IMEMode;  // Returns or sets the default start-up mode for the Japanese Input Method Editor, IME
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
@property WordE198 windowState;  // Returns or sets the state of the window.
@property (readonly) WordE201 windowType;  // Returns the window type.

- (void) toggleShowAllReviewers;  // Toggles whether or not all reviewers are shown in this window.

@end

// Contains magnification options, for example, the zoom percentage for a window or pane.
@interface WordZoom : WordBaseObject

@property NSInteger pageColumns;  // Returns or sets the number of pages to be displayed side by side on-screen at the same time in print layout view or print preview.
@property WordE205 pageFit;  // Returns or sets the view magnification of a window so that either the entire page is visible or the entire width of the page is visible.
@property NSInteger pageRows;  // Returns or sets the number of pages to be displayed one above the other on-screen at the same time in print layout view or print preview. 
@property NSInteger percentage;  // Returns or sets the magnification for a window as a percentage.


@end



/*
 * Drawing Suite
 */

@interface WordAdjustment : WordBaseObject

@property double adjustment_value;  // Returns or sets a floating point adjustment for a shape.


@end

// Contains properties and methods that apply to line callouts.
@interface WordCalloutFormat : WordBaseObject

@property BOOL accent;  // Returns or sets if a vertical accent bar separates the callout text from the callout line.
@property WordMCAt angle;  // Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box.
@property BOOL autoAttach;  // Returns or sets if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line, where the callout points to, is to the left or right of the callout text box.
@property (readonly) BOOL autoLength;  // Returns if the length of the callout line is automatically set. Use the automatic length method to set this property to true, and use the custom length method to set this property to false.
@property (readonly) double calloutFormatLength;  // When the AutoLength property of the specified callout is set to False, the Length property returns the length in points of the first segment of the callout line, the segment attached to the text callout box.
@property BOOL calloutHasBorder;  // Returns or sets whether the text in the specified callout is surrounded by a border.
@property WordMCot calloutType;  // Returns or sets the callout type.
@property (readonly) double drop;  // For callouts with an explicitly set drop value, this property returns the vertical distance in points from the edge of the text bounding box to the place where the callout line attaches to the text box.
@property (readonly) WordMCDt dropType;  // Returns a value that indicates where the callout line attaches to the callout text box.
@property double gap;  // Returns or sets the horizontal distance in points between the end of the callout line and the text bounding box.


@end

// Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.
@interface WordFillFormat : WordBaseObject

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property WordMCoI backColorThemeIndex;  // Returns or sets the specified fill background color.
@property (readonly) WordMFdT fillType;  // Returns the shape fill format type.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property WordMCoI foreColorThemeIndex;  // Returns or sets the specified foreground fill or solid color.
@property (readonly) WordMGCt gradientColorType;  // Returns the gradient color type for the specified fill.
@property (readonly) double gradientDegree;  // Returns a value that indicates how dark or light a one-color gradient fill is. A value of zero means that black is mixed in with the shape's foreground color to form the gradient; a value of 1 means that white is mixed in. Values between 1 and zero blend.
@property (readonly) WordMGdS gradientStyle;  // Returns the gradient style for the specified fill.
@property (readonly) NSInteger gradientVariant;  // Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. If the gradient style is from center gradient, this property returns either 1 or 2.
@property (readonly) WordPpTy pattern;  // Returns the value that represents the pattern applied to the specified fill or line.
@property (readonly) WordMPGb presetGradientType;  // Returns the preset gradient type for the specified fill.
@property (readonly) WordMPzT presetTexture;  // Returns the preset texture for the specified fill.
@property (copy, readonly) NSString *textureName;  // Returns the name of the custom texture file for the specified fill.
@property (readonly) WordMxtT textureType;  // Returns the texture type for the specified fill.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque, and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Represents the glow formatting for a shape or range of shapes.
@interface WordGlowFormat : WordBaseObject

@property (copy) NSColor *color;  // Returns or sets the color for the specified glow format.
@property WordMCoI colorThemeIndex;  // Returns or sets the color for the specified glow format.
@property double radius;  // Returns or sets the length of the radius for the specified glow format.


@end

// Represents horizontal line formatting.
@interface WordHorizontalLineFormat : WordBaseObject

@property WordE280 alignment;  // Returns or sets a constant that represents the alignment for the specified horizontal line.
@property BOOL noShade;  // Returns or sets if Microsoft Word draws the specified horizontal line without 3-D shading.
@property double percentWidth;  // Returns or sets the length of the specified horizontal line expressed as a percentage of the window width.
@property WordE281 widthType;  // Returns or sets the width type for the specified horizontal line format object.


@end

// Represents a graphic object in the text layer of a document.
@interface WordInlineShape : WordBaseObject

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
@property (readonly) WordE259 inlineShapeType;  // The type of this inline shape.
@property BOOL isInlinePlaceholder;  // Returns true if a shape is a placeholder.
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
- (WordBorder *) getBorderWhichBorder:(WordE122)whichBorder;  // Returns the specified border object.
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
@interface WordLineFormat : WordBaseObject

@property (copy) NSColor *backColor;  // Returns or sets a RGB color that represents the background color for the specified fill or patterned line.
@property WordMCoI backColorThemeIndex;  // Returns or sets the background color for a patterned line.
@property WordMAhL beginArrowheadLength;  // Returns or sets the length of the arrowhead at the beginning of the specified line.
@property WordMAhS beginArrowheadStyle;  // Returns or sets the style of the arrowhead at the beginning of the specified line.
@property WordMAhW beginArrowheadWidth;  // Returns or sets the width of the arrowhead at the beginning of the specified line.
@property WordMlDs dashStyle;  // Returns or sets the dash style for the specified line.
@property WordMAhL endArrowheadLength;  // Returns or sets the length of the arrowhead at the end of the specified line.
@property WordMAhS endArrowheadStyle;  // Returns or sets the style of the arrowhead at the end of the specified line.
@property WordMAhW endArrowheadWidth;  // Returns or sets the width of the arrowhead at the end of the specified line.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property WordMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the line.
@property WordMLnS lineStyle;  // Returns or sets the line format style.
@property WordPpTy pattern;  // Returns or sets a value that represents the pattern applied to the specified fill or line.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.
@property double weight;  // Returns or sets the thickness of the specified line in points.


@end

// Represents a Microsoft Office theme.
@interface WordOfficeTheme : WordBaseObject

@property (copy, readonly) WordThemeColorScheme *themeColorScheme;  // Returns the color scheme of a Microsoft Office theme.
@property (copy, readonly) WordThemeEffectScheme *themeEffectScheme;  // Returns the effects scheme of a Microsoft Office theme.
@property (copy, readonly) WordThemeFontScheme *themeFontScheme;  // Returns the font scheme of a Microsoft Office theme.


@end

// Contains properties and methods that apply to pictures. 
@interface WordPictureFormat : WordBaseObject

@property double brightness;  // Returns or sets the brightness of the specified picture . The value for this property must be a number from 0.0, dimmest to 1.0, brightest.
@property double contrast;  // Returns or sets the contrast for the specified picture. The value for this property must be a number from 0.0, the least contrast to 1.0, the greatest contrast.
@property double cropBottom;  // Returns or sets the number of points that are cropped off the bottom of the specified picture.
@property double cropLeft;  // Returns or sets the number of points that are cropped off the left side of the specified picture.
@property double cropRight;  // Returns or sets the number of points that are cropped off the right side of the specified picture.
@property double cropTop;  // Returns or sets the number of points that are cropped off the top of the specified picture.
@property (copy) NSColor *transparencyColor;  // Returns or sets the transparent color for the specified picture as a RGB color. For this property to take effect, the transparent background property must be set to true.
@property BOOL transparentBackground;  // Returns or sets if the parts of the picture that are defined with a transparent color actually appear transparent.


@end

// Represents the reflection effect in Office graphics.
@interface WordReflectionFormat : WordBaseObject

@property WordMRfT reflectionType;  // Returns or sets the type of the reflection format object.


@end

// Represents shadow formatting for a shape.
@interface WordShadowFormat : WordBaseObject

@property double blur;  // Returns or sets the blur, in points, of the specified shadow.
@property (copy) NSColor *foreColor;  // Returns or sets a RGB color that represents the foreground color for the fill, line, or shadow.
@property WordMCoI foreColorThemeIndex;  // Returns or sets the foreground color for the shadow format.
@property BOOL obscured;  // Returns or sets if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. If false the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
@property double offsetX;  // Returns or sets the horizontal offset in points of the shadow from the specified shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
@property double offsetY;  // Returns or sets the vertical offset in points of the shadow from the specified shape. A positive value offsets the shadow below the shape; a negative value offsets it above the shape.
@property BOOL rotateWithShape;  // Returns or sets whether to rotate the shadow when rotating the shape.
@property WordMSSt shadowStyle;  // Returns or sets the style of shadow formatting to apply to a shape.
@property WordMSdT shadowType;  // Returns or sets the shape shadow type.
@property double size;  // Returns or sets the width of the shadow.
@property double transparency;  // Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0, opaque and 1.0, clear.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Represents an object in the drawing layer.
@interface WordShape : WordBaseObject

- (SBElementArray *) adjustments;

@property (copy, readonly) WordTextRange *anchor;  // Returns a text range object that represents the anchoring range for the specified shape.
@property (readonly) NSInteger anchorID;  // Returns the anchor id for the specified shape.
@property WordMAsT autoShapeType;  // Returns or sets the shape type for the specified shape object, which must represent an autoshape.
@property (readonly) BOOL child;  // True if the shape is a child shape.
@property (readonly) NSInteger editID;  // Returns the edit id for the specified shape.
@property (copy, readonly) WordFillFormat *fillFormat;  // Return the fill format object associated with this shape object.
@property (copy, readonly) WordGlowFormat *glowFormat;  // Returns the formatting properties for a glow effect.
@property BOOL hasChart;  // True if the specified shape has a chart.
@property double height;  // Returns or sets the height of the object.
@property (readonly) BOOL horizontalFlip;  // Returns true if the shape has been flipped horizontally. 
@property (copy, readonly) WordHyperlinkObject *hyperlink;  // Returns the hyperlink object associated with this shape object.
@property BOOL isPlaceholder;  // Returns true if a shape is a placeholder.
@property double leftPosition;  // Returns or sets the horizontal position of the object.
@property (copy, readonly) WordLineFormat *lineFormat;  // Returns the line format object associated with this shape object.
@property (copy, readonly) WordLinkFormat *linkFormat;  // Returns the link format object associated with this shape object.
@property BOOL lockAnchor;  // Returns or sets if the specified shape object's anchor is locked to the anchoring range. When a shape has a locked anchor, you cannot move the shape's anchor by dragging it, the anchor doesn't move as the shape is moved.
@property BOOL lockAspectRatio;  // Returns or sets if the specified shape retains its original proportions when you resize it. If false, you can change the height and width of the shape independently of one another when you resize it.
@property (copy) NSString *name;  // Returns or sets the name of this shape object.
@property (copy, readonly) WordReflectionFormat *reflectionFormat;  // Returns the formatting properties for a reflection effect.
@property WordE236 relativeHorizontalPosition;  // Returns or sets if the horizontal position of the shape is relative.
@property WordE237 relativeVerticalPosition;  // Returns or sets if the vertical position of the shape is relative.
@property double rotation;  // Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation.
@property (copy, readonly) WordShadowFormat *shadow;  // Returns the shadow format object associated with this shape object.
@property (readonly) WordMShp shapeType;  // Returns the shape type.
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
- (void) flipFlipCommand:(WordMFlC)flipCommand;  // Flips a shape horizontally or vertically.
- (void) pickUp;  // Copies the formatting of the specified shape. Use the apply method to apply the copied formatting to another shape.
- (void) rerouteConnections;  // Reroutes the connections between shapes.
- (void) setShapesDefaultProperties;  // Applies the formatting of the specified shape to a default shape for that document. New shapes inherit many of their attributes from the default shape.
- (void) zOrderZOrderCommand:(WordMZoC)zOrderCommand;  // Moves the specified shape in front of or behind other shapes.

@end

// Represents an callout object in the drawing layer.
@interface WordCallout : WordShape

@property (copy, readonly) WordCalloutFormat *calloutFormat;  // Returns the callout format object associated with this callout shape object.
@property (readonly) WordMCot calloutType;  // Return the type of this callout


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

- (void) scaleHeightFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(WordMSFr)scale;  // Scales the height of the picture shape by a specified factor.
- (void) scaleWidthFactor:(double)factor relativeToOriginalSize:(BOOL)relativeToOriginalSize scale:(WordMSFr)scale;  // Scales the width of the shape by a specified factor.

@end

// Represents the soft edge formatting for a shape or range of shapes.
@interface WordSoftEdgeFormat : WordBaseObject

@property WordMSET softEdgeType;  // Returns or sets the type soft edge format object.


@end

// Represents a standard horizontal line object in the text layer of a document.
@interface WordStandardInlineHorizontalLine : WordInlineShape


@end

// Represents an text box object in the drawing layer.
@interface WordTextBox : WordShape

@property (readonly) WordTxOr textOrientation;  // Returns the orientation of the text inside the text shape.


@end

// Represents the text frame in a shape object. Contains the text in the text frame as well as the properties that control the margins and orientation of the text frame.
@interface WordTextFrame : WordBaseObject

@property (copy, readonly) WordTextRange *containingRange;  // Returns a text range object that represents the entire story in a series of shapes with linked text frames that the specified text frame belongs to.
@property (readonly) BOOL hasText;  // Returns true if the specified shape has text associated with it.
@property double marginBottom;  // Returns or sets the distance in points between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text.
@property double marginLeft;  // Returns or sets the distance in points between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text.
@property double marginRight;  // Returns or sets the distance in points between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text.
@property double marginTop;  // Returns or sets the distance in points between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text.
@property (copy) WordTextFrame *nextTextframe;  // Returns or sets the next text frame object.
@property WordTxOr orientation;  // Returns or sets the orientation of the text inside the frame.
@property (readonly) BOOL overflowing;  // Returns if the text inside the specified text frame doesn't all fit within the frame.
@property (copy) WordTextFrame *previousTextframe;  // Returns or sets the previous text frame object.
@property (copy, readonly) WordTextRange *textRange;  // Returns a text range object that represents the text range for this text frame.
@property BOOL textwrappingAllowed;  // Allow textwrapping of overlay objects.

- (void) breakForwardLink;  // Breaks the forward link for the specified text frame, if such a link exists.
- (BOOL) validLinkTargetTargetTextframe:(WordTextFrame *)targetTextframe;  // Determines whether the text frame of one shape can be linked to the text frame of another shape. Returns false if target textframe already contains text or is already linked, or if the shape doesn't support attached text.

@end

// Represents the color scheme of a Microsoft Office 2007 theme.
@interface WordThemeColorScheme : WordBaseObject

- (SBElementArray *) themeColors;

- (NSColor *) getCustomColorName:(NSString *)name;  // Returns the custom color for the specified Microsoft Office theme.
- (void) loadThemeColorSchemeFileName:(NSString *)fileName;  // Loads the color scheme of a Microsoft Office theme from a file
- (void) saveThemeColorSchemeFileName:(NSString *)fileName;  // Saves the color scheme of a Microsoft Office theme to a file.

@end

// Represents a color in the color scheme of a Microsoft Office 2007 theme.
@interface WordThemeColor : WordBaseObject

@property (copy) NSColor *RGB;  // Returns or sets a value of a color in the color scheme of a Microsoft Office theme.
@property (readonly) WordMCSI themeColorSchemeIndex;  // Returns the index value a color scheme of a Microsoft Office theme.


@end

// Represents the effect scheme of a Microsoft Office theme.
@interface WordThemeEffectScheme : WordBaseObject

- (void) loadThemeEffectSchemeFileName:(NSString *)fileName;  // Loads the effects scheme of a Microsoft Office theme from a file

@end

// Represents the font scheme of a Microsoft Office theme.
@interface WordThemeFontScheme : WordBaseObject

- (SBElementArray *) minorThemeFonts;
- (SBElementArray *) majorThemeFonts;

- (void) loadThemeFontSchemeFileName:(NSString *)fileName;  // Loads the font scheme of a Microsoft Office theme from a file.
- (void) saveThemeFontSchemeFileName:(NSString *)fileName;  // Saves the font scheme of a Microsoft Office theme to a file.

@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface WordThemeFont : WordBaseObject

@property (copy) NSString *name;  // Returns or sets a value specifying the font to use for a selection.


@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface WordMajorThemeFont : WordThemeFont


@end

// Represents a container for the font schemes of a Microsoft Office 2007 theme.
@interface WordMinorThemeFont : WordThemeFont


@end

// Represents a collection of major and minor fonts in the font scheme of a Microsoft Office 2007 theme.
@interface WordThemeFonts : WordBaseObject


@end

// Represents a shape's three-dimensional formatting.
@interface WordThreeDFormat : WordBaseObject

@property double ZDistance;  // Returns or sets a Single that represents the distance from the center of an object or text.
@property double bevelBottomDepth;  // Returns or sets the depth/height of the bottom bevel.
@property double bevelBottomInset;  // Returns or sets the inset size/width for the bottom bevel.
@property WordMBlT bevelBottomType;  // Returns or sets the bevel type for the bottom bevel.
@property double bevelTopDepth;  // Returns or sets the depth/height of the top bevel.
@property double bevelTopInset;  // Returns or sets the inset size/width for the top bevel.
@property WordMBlT bevelTopType;  // Returns or sets the bevel type for the top bevel.
@property (copy) NSColor *contourColor;  // Returns or sets the color of the contour of an object or text.
@property WordMCoI contourColorThemeIndex;  // Returns or sets the color for the specified contour.
@property double contourWidth;  // Returns or sets the width of the contour of an object or text.
@property double depth;  // Returns or sets the depth of the shape's extrusion.
@property (copy) NSColor *extrusionColor;  // Returns or sets a RGB color that represents the color of the shape's extrusion.
@property WordMCoI extrusionColorThemeIndex;  // Returns or sets the color for the specified extrusion.
@property WordMExC extrusionColorType;  // Returns or sets a value that indicates what will determine the extrusion color.
@property double fieldOfView;  // Returns or sets the amount of perspective for an object or text.
@property double lightAngle;  // Returns or sets the angle of the lighting.
@property BOOL perspective;  // Returns or sets if the extrusion appears in perspective that is, if the walls of the extrusion narrow toward a vanishing point. If false, the extrusion is a parallel, or orthographic, projection that is, if the walls don't narrow toward a vanishing point.
@property WordMPzC presetCamera;  // Returns a constant that represents the camera preset.
@property WordMExD presetExtrusionDirection;  // Returns or sets the direction taken by the extrusion's sweep path leading away from the extruded shape, the front face of the extrusion.
@property WordMPLd presetLightingDirection;  // Returns or sets the position of the light source relative to the extrusion.
@property WordMLtT presetLightingRig;  // Returns a constant that represents the lighting preset.
@property WordMlSf presetLightingSoftness;  // Returns or sets the intensity of the extrusion lighting.
@property WordMPMt presetMaterial;  // Returns or sets the extrusion surface material.
@property WordM3DF presetThreeDFormat;  // Returns or sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.
@property BOOL projectText;  // Returns or sets whether text on a shape rotates with shape.
@property double rotationX;  // Returns or sets the rotation of the extruded shape around the x-axis in degrees. A positive value indicates upward rotation; a negative value indicates downward rotation.
@property double rotationY;  // Returns or sets the rotation of the extruded shape around the y-axis, in degrees. A positive value indicates rotation to the left; a negative value indicates rotation to the right.
@property double rotationZ;  // Returns or sets the rotation of the extruded shape around the z-axis, in degrees. A positive value indicates clockwise rotation; a negative value indicates anti-clockwise rotation.
@property BOOL visible;  // Returns or sets if the specified object, or the formatting applied to it, is visible.


@end

// Contains properties and methods that apply to WordArt objects.
@interface WordWordArtFormat : WordBaseObject

@property WordMTxA alignment;  // Returns or sets a constant that represents the alignment for the specified text effect.
@property BOOL bold;  // Returns or sets if the text of the word art shape is formatted as bold.
@property (copy) NSString *fontName;  // Returns or sets the font name of the font used by this word art shape.
@property double fontSize;  // Returns or sets the font size of the font used by this word art shape.
@property BOOL italic;  // Returns or sets if the text of the word art shape is formatted as italic.
@property BOOL kernedPairs;  // Returns or sets if character pairs in a WordArt object have been kerned. 
@property BOOL normalizedHeight;  // Returns or sets if all characters, both uppercase and lowercase, in the specified WordArt are the same height.
@property WordMPTs presetShape;  // Returns or sets the shape of the specified WordArt.
@property WordMPXF presetWordArtEffect;  // Returns or sets the style of the specified WordArt.
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
@property (readonly) WordMPXF presetWordArtEffect;  // Returns the style of the specified word art.
@property (copy, readonly) WordWordArtFormat *wordArtFormat;  // Returns the word art format object associated with the word art shape object.
@property (copy, readonly) NSString *wordArtText;  // The text associated with this word art shape.


@end

// Represents all the properties for wrapping text around a shape.
@interface WordWrapFormat : WordBaseObject

@property BOOL allowOverlap;  // Returns or sets a value that specifies whether a given shape can overlap other shapes.
@property double distanceBottom;  // Returns or sets the distance in points between the document text and the bottom edge of the text-free area surrounding the specified shape.
@property double distanceLeft;  // Returns or sets the distance in points between the document text and the left edge of the text-free area surrounding the specified shape.
@property double distanceRight;  // Returns or sets the distance in points between the document text and the right edge of the text-free area surrounding the specified shape.
@property double distanceTop;  // Returns or sets the distance in points between the document text and the top edge of the text-free area surrounding the specified shape.
@property WordE268 wrapSide;  // Returns or sets a value that indicates whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.
@property WordE267 wrapType;  // Returns or sets the wrap type for the specified shape.


@end



/*
 * Text Suite
 */

// Represents a single built-in or user-defined style. The Word style object includes style attributes, font, font style, paragraph spacing, and so on, as properties of the Word style object.
@interface WordWordStyle : WordBaseObject

@property BOOL automaticallyUpdate;  // True if the style is automatically redefined based on the selection. False if Word prompts for confirmation before redefining the style based on the selection.
@property WordE184 baseStyle;  // Returns or sets an existing style on which you can base the formatting of another style.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property (readonly) BOOL builtIn;  // Returns true if the specified object is one of the built-in styles in Word.
@property (copy, readonly) NSString *objectDescription;  // Returns the description of the specified style. For example, a typical description for the Heading 2 style might be Normal, Font: Arial, 12 pt, Bold, Italic, Space Before 12 pt After 3 pt, KeepWithNext, Level 2.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this Word style object.
@property (copy, readonly) WordFrame *frame;  // Returns the frame object associated with this Word style object.
@property (readonly) BOOL inUse;  // Returns true if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.
@property WordE182 languageID;  // Returns or sets the language for the Word style object
@property WordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (readonly) NSInteger listLevelNumber;  // Returns the list level for the specified style.
@property (copy, readonly) WordListTemplate *listTemplate;  // Returns the list template object associated with this Word style object.
@property (copy) NSString *nameLocal;  // Returns or sets the name of a built-in style in the language of the user. Setting this property renames a user-defined style or adds an alias to a built-in style. 
@property WordE184 nextParagraphStyle;  // Returns or sets the style to be applied automatically to a new paragraph inserted after a paragraph formatted with the specified style.
@property BOOL noProofing;  // Returns or sets if Microsoft Word finds or replaces text that the spelling and grammar checker ignores for this Word style
@property BOOL noSpaceBetweenSame;  // Returns or sets if Microsoft Word suppresses space between paragraphs of the same style
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this Word style object.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the Word style.
@property (readonly) WordE128 styleType;  // Returns the style type.
@property (copy, readonly) WordTableStyle *tableStyle;  // Returns table style properties for this style.

- (void) linkToListTemplateListTemplate:(WordListTemplate *)listTemplate listLevelNumber:(NSInteger)listLevelNumber;  // Links the specified style to a list template so that the style's formatting can be applied to lists.

@end

// Represents all the formatting for a paragraph.
@interface WordParagraphFormat : WordBaseObject

- (SBElementArray *) tabStops;

@property BOOL addSpaceBetweenEastAsianAndAlpha;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs.
@property BOOL addSpaceBetweenEastAsianAndDigit;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs.
@property WordE142 alignment;  // Returns or sets a constant that represents the alignment for the specified paragraphs.
@property BOOL autoAdjustRightIndent;  // Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you’ve specified a set number of characters per line.
@property WordE104 baseLineAlignment;  // Returns or sets a constant that represents the vertical position of fonts on a line.
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
@property WordE157 lineSpacingRule;  // Returns or sets the line spacing for the specified paragraphs.
@property double lineUnitAfter;  // Returns or sets the amount of spacing in gridlines after the specified paragraphs.
@property double lineUnitBefore;  // Returns or sets the amount of spacing in gridlines before the specified paragraphs.
@property BOOL noLineNumber;  // Returns or set if line numbers are repressed for the specified paragraphs.
@property WordE269 outlineLevel;  // Returns or sets the outline level for the specified paragraphs.
@property BOOL pageBreakBefore;  // Returns or sets if a page break is forced before the specified paragraphs.
@property double paragraphFormatLeftIndent;  // Returns or sets the left indent in points for the specified paragraphs.
@property double paragraphFormatRightIndent;  // Returns or sets the right indent in points for the specified paragraphs.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the paragraph format object.
@property double spaceAfter;  // Returns or sets the spacing in points after the specified paragraphs.
@property BOOL spaceAfterAuto;  // Returns or sets if Microsoft Word automatically sets the amount of spacing after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing in points before the specified paragraphs.
@property BOOL spaceBeforeAuto;  // Returns or sets if Microsoft Word automatically sets the amount of spacing before the specified paragraphs.
@property WordE184 style;  // Returns or sets the Word style associated with the replacement object.
@property BOOL widowControl;  // Returns or sets if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document.
@property BOOL wordWrap;  // Returns or sets if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames.


@end

// Represents a single paragraph in a selection, range, or document.
@interface WordParagraph : WordBaseObject

- (SBElementArray *) tabStops;

@property BOOL addSpaceBetweenEastAsianAndAlpha;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs.
@property BOOL addSpaceBetweenEastAsianAndDigit;  // Returns or sets if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs.
@property WordE142 alignment;  // Returns or sets a constant that represents the alignment for the specified paragraphs.
@property BOOL autoAdjustRightIndent;  // Returns or sets if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you’ve specified a set number of characters per line.
@property WordE104 baseLineAlignment;  // Returns or sets a constant that represents the vertical position of fonts on a line.
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
@property WordE157 lineSpacingRule;  // Returns or sets the line spacing for the specified paragraphs.
@property double lineUnitAfter;  // Returns or sets the amount of spacing in gridlines after the specified paragraphs.
@property double lineUnitBefore;  // Returns or sets the amount of spacing in gridlines before the specified paragraphs.
@property BOOL noLineNumber;  // Returns or set if line numbers are repressed for the specified paragraphs.
@property WordE269 outlineLevel;  // Returns or sets the outline level for the specified paragraphs.
@property BOOL pageBreakBefore;  // Returns or sets if a page break is forced before the specified paragraphs.
@property (copy) WordParagraphFormat *paragraphFormat;  // Returns or sets the paragraph format object associated with this paragraph object.
@property double paragraphLeftIndent;  // Returns or sets the left indent in points for the specified paragraphs.
@property NSInteger paragraph_id;  // Returns the paragraph id for the specified paragraphs.
@property double rightIndent;  // Returns or sets the right indent in points for the specified paragraphs.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the paragraph object.
@property double spaceAfter;  // Returns or sets the spacing in points after the specified paragraphs.
@property double spaceBefore;  // Returns or sets the spacing in points before the specified paragraphs.
@property WordE184 style;  // Returns or sets the Word style associated with the replacement object.
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
@interface WordSection : WordBaseObject

@property (copy, readonly) WordBorderOptions *borderOptions;
@property (copy) WordPageSetup *pageSetup;  // Returns or sets a page setup object associated with this section object
@property BOOL protectedForForms;  // Returns or sets if the specified section is protected for forms. When a section is protected for forms, you can select and modify text only in form fields.
@property (readonly) NSInteger sectionIndex;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the section object.

- (WordHeaderFooter *) getFooterIndex:(WordE163)index;  // Returns a specific footer object
- (WordHeaderFooter *) getHeaderIndex:(WordE163)index;  // Returns a specific header object

@end

// Contains shading attributes for an object.
@interface WordShading : WordBaseObject

@property (copy) NSColor *backgroundPatternColor;  // Returns or sets the RGB color that's applied to the background of the shading object.
@property WordE110 backgroundPatternColorIndex;  // Returns or sets the color index that's applied to the background of the shading object.
@property WordMCoI backgroundPatternColorThemeIndex;  // Returns or sets the color that's applied to the background of the shading object.
@property (copy) NSColor *foregroundPatternColor;  // Returns or sets the RGB color that's applied to the foreground of the shading object.
@property WordE110 foregroundPatternColorIndex;  // Returns or sets the color index that's applied to the foreground of the shading object.
@property WordMCoI foregroundPatternColorThemeIndex;  // Returns or sets the color that's applied to the foreground of the shading object. This color is applied to the dots and lines in the shading pattern.
@property WordE112 texture;  // Returns or sets the shading texture for the specified object.


@end

// Represents a contiguous area in a document. Each text range object is defined by a starting and ending character position.
@interface WordTextRange : WordBaseObject

- (SBElementArray *) characters;
- (SBElementArray *) words;
- (SBElementArray *) sentences;
- (SBElementArray *) tables;
- (SBElementArray *) footnotes;
- (SBElementArray *) endnotes;
- (SBElementArray *) WordComments;
- (SBElementArray *) cells;
- (SBElementArray *) sections;
- (SBElementArray *) paragraphs;
- (SBElementArray *) fields;
- (SBElementArray *) formFields;
- (SBElementArray *) frames;
- (SBElementArray *) bookmarks;
- (SBElementArray *) revisions;
- (SBElementArray *) hyperlinkObjects;
- (SBElementArray *) subdocuments;
- (SBElementArray *) columns;
- (SBElementArray *) rows;
- (SBElementArray *) shapes;
- (SBElementArray *) readabilityStatistics;
- (SBElementArray *) grammaticalErrors;
- (SBElementArray *) spellingErrors;
- (SBElementArray *) inlineShapes;
- (SBElementArray *) mathObjects;
- (SBElementArray *) coauthLocks;
- (SBElementArray *) coauthUpdates;
- (SBElementArray *) conflicts;

@property BOOL bold;  // Returns or sets if the text associated with the text range is formatted as bold.
@property (readonly) NSInteger bookmarkId;  // Returns the number of the bookmark that encloses the beginning of the text range. The number corresponds to the position of the bookmark in the document, 1 for the first bookmark, 2 for the second one, and so on.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this text range object.
@property WordE125 character_case;  // Synonym for case
- (WordE125) character_case;  // Returns or sets a constant that represents the case of the text in the text range.
- (void) setCase: (WordE125) character_case;
@property (copy, readonly) WordColumnOptions *columnOptions;
@property BOOL combineCharacters;
@property (copy) NSString *content;  // Returns or sets the text in the text range.
@property BOOL disableCharacterSpaceGrid;  // Returns or sets if Microsoft Word ignores the number of characters per line for the corresponding text range object.
@property WordE114 emphasisMark;  // Returns or sets the emphasis mark for a character or designated character string.
@property (readonly) NSInteger endOfContent;  // Returns the ending character position of the text range.
@property (copy, readonly) WordEndnoteOptions *endnoteOptions;  // Returns the endnote options object for this text range.
@property (copy, readonly) WordFieldOptions *fieldOptions;
@property (copy, readonly) WordFind *findObject;  // Returns the find object associated with this text range.
@property double fitTextWidth;  // Returns or sets the width in which Microsoft Word fits the text in the text range.
@property (copy, readonly) WordFont *fontObject;  // Returns the font object associated with this text range.
@property (copy, readonly) WordFootnoteOptions *footnoteOptions;  // Returns the footnote options object for this text range.
@property (copy) WordTextRange *formattedText;  // Returns or sets a text range object that includes the formatted text in the text range.
@property BOOL grammarChecked;  // True if a grammar check has been run on the text range. False if some of the text range hasn't been checked for grammar. To recheck the grammar in the document, set the grammar checked property to false.
@property WordE110 highlightColorIndex;  // Returns or sets the highlight color for the text range.
@property WordE309 horizontalInVertical;
@property (readonly) BOOL isEndOfRowMark;  // Returns true if the text range is collapsed and is located at the end-of-row mark in a table.
@property BOOL italic;  // Returns or sets if the text associated with the text range is formatted as italic.
@property WordE182 languageID;  // Returns or sets the language for the text range object
@property WordE182 languageIDEastAsian;  // Returns or sets an east asian language for the template.
@property (copy, readonly) WordListFormat *listFormat;  // Returns the list format object associated with this text range.
@property (copy, readonly) WordTextRange *nextStoryRange;  // Returns a text range object that refers to the next story
@property BOOL noProofing;  // Returns or sets if the spelling and grammar checker should ignore documents based on this text range.
@property WordE270 orientation;  // Returns or sets the orientation of text in a text range when the text direction feature is enabled.
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
@property (readonly) WordE160 storyType;  // Returns the story type for the text range.
@property WordE184 style;  // Returns or sets the Word style for this range.
@property BOOL subdocumentsExpanded;  // Returns or sets if the subdocuments in the specified text range are expanded.
@property WordE182 supplementalLanguageID;  // Returns or sets the language for the text range object
@property (copy) WordTextRetrievalMode *textRetrievalMode;  // Returns or sets the text retrieval object that controls how text is retrieved from this text range.
@property WordE308 twoLinesInOne;  // Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.  Read/write
@property WordE113 underline;  // Returns or sets the type of underline applied to the text range.

- (void) autoFormatTextRange;  // Automatically formats a text range.
- (double) calculateRange;  // Calculates a mathematical expression within a text range.
- (WordTextRange *) changeEndOfRangeBy:(WordE129)by extendType:(WordE249)extendType;  // Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) changeStartOfRangeBy:(WordE129)by extendType:(WordE249)extendType;  // Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns the new range object or missing value if an error occurred.
- (BOOL) checkGrammar;  // Begins a grammar check for the text range.  Returns true if the text range had no errors
- (BOOL) checkSpellingCustomDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Begins a spelling check for the text range.  Returns true if the text range had no errors
- (void) checkSynonyms;  // Displays the thesaurus dialog box, which lists alternative word choices, or synonyms, for the text in the text range.
- (WordTextRange *) collapseRangeDirection:(WordE132)direction;  // Collapses this text range to the starting or ending position and returns a new text range object. After a text range is collapsed, the starting and ending points are equal.
- (NSInteger) computeTextRangeStatisticsStatistic:(WordE155)statistic;  // Returns a statistic based on the contents of the specified text range.
- (WordTable *) convertToTableSeparator:(WordE177)separator numberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns initialColumnWidth:(NSInteger)initialColumnWidth tableFormat:(WordE180)tableFormat applyBorders:(BOOL)applyBorders applyShading:(BOOL)applyShading applyFont:(BOOL)applyFont applyColor:(BOOL)applyColor applyHeadingRows:(BOOL)applyHeadingRows applyLastRow:(BOOL)applyLastRow applyFirstColumn:(BOOL)applyFirstColumn applyLastColumn:(BOOL)applyLastColumn autoFit:(BOOL)autoFit autoFitBehavior:(WordE288)autoFitBehavior defaultTableBehavior:(WordE287)defaultTableBehavior;  // Converts text within a text range to a table.
- (void) copyAsPicture;  // Copies the content of the text range as a picture.
- (void) copyObject;  // Copies the content of the text range to the clipboard.
- (void) cutObject;  // Removes the content of the text range from the document and places it on the clipboard.
- (WordTextRange *) expandBy:(WordE129)by;  // Expands the specified range. This method returns the new range object or missing value if an error occurred.
- (NSString *) getRangeInformationInformationType:(WordE266)informationType;  // Returns requested information about the text range. 
- (WordTextRange *) goToNextWhat:(WordE130)what;  // Returns a text range object that refers to the start position of the next item or location specified by the what argument.
- (WordTextRange *) goToPreviousWhat:(WordE130)what;  // Returns a text range object that refers to the start position of the previous item or location specified by the what argument.
- (BOOL) inRangeTextRange:(WordTextRange *)textRange;  // Returns true if the text range to which the method is applied is contained in the range specified by the text range argument.
- (BOOL) inStoryTextRange:(WordTextRange *)textRange;  // Returns true if the text range to which the method is applied is in the same story as the range specified by the text range argument.
- (BOOL) isEquivalentTextRange:(WordTextRange *)textRange;  // True if the selection or range to which this method is applied is equal to the range specified by the text range argument. This method compares the starting and ending character positions, as well as the story type.
- (void) modifyEnclosureEnclosureStyle:(WordE277)enclosureStyle enclosureType:(WordE276)enclosureType enclosedText:(NSString *)enclosedText;  // Adds, modifies, or removes an enclosure around the specified character or characters.
- (WordTextRange *) moveEndOfRangeBy:(WordE129)by count:(NSInteger)count;  // Moves the ending character position of the range.  This method returns the new text range object or missing value if an error occurred.
- (WordTextRange *) moveRangeBy:(WordE129)by count:(NSInteger)count;  // Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeEndUntilCharacters:(NSString *)characters count:(WordE294)count;  // Moves the end position of the specified range until any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeEndWhileCharacters:(NSString *)characters count:(WordE294)count;  // Moves the ending character position of a the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeStartUntilCharacters:(NSString *)characters count:(WordE294)count;  // Moves the start position of the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeStartWhileCharacters:(NSString *)characters count:(WordE294)count;  // Moves the start position of the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeUntilCharacters:(NSString *)characters count:(WordE294)count;  // Moves the specified range until one of the specified characters is found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveRangeWhileCharacters:(NSString *)characters count:(WordE294)count;  // Moves the specified range while any of the specified characters are found in the document. This method returns the new range object or missing value if an error occurred.
- (WordTextRange *) moveStartOfRangeBy:(WordE129)by count:(NSInteger)count;  // Moves the starting character position of the range. This method returns the new text range object or missing value if an error occurred.
- (WordTextRange *) navigateTo:(WordE130)to position:(WordE131)position count:(NSInteger)count name:(NSString *)name;  // Returns a text range object that represents the start position of the specified item, such as a page, bookmark, or field.
- (WordTextRange *) nextRangeBy:(WordE129)by count:(NSInteger)count;  // Returns a text range object relative to the specified text range.
- (WordTextRange *) nextSubdocument;  // Returns a new text range object to the next subdocument. If there isn't another subdocument, an error occurs.
- (void) pasteAndFormatType:(WordE293)type;  // Pastes the selected table cells and formats them as specified.
- (void) pasteAppendTable;  // Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.
- (void) pasteAsNestedTable;  // Pastes a cell or group of cells as a nested table into the text range.
- (void) pasteExcelTableLinkedToExcel:(BOOL)linkedToExcel wordFormatting:(BOOL)wordFormatting RTF:(BOOL)RTF;  // Pastes and formats a Microsoft Excel table.
- (void) pasteObject;  // Inserts the contents of the clipboard at the specified text range.
- (void) pasteSpecialLink:(BOOL)link placement:(WordE243)placement displayAsIcon:(BOOL)displayAsIcon dataType:(WordE251)dataType iconLabel:(NSString *)iconLabel;  // Inserts the contents of the clipboard. Unlike with the paste method, with paste special you can control the format of the pasted information and optionally establish a link to the source file - for example, a Microsoft Excel worksheet.
- (void) phoneticGuideText:(NSString *)text alignmentType:(WordE310)alignmentType raise:(NSInteger)raise fontSize:(NSInteger)fontSize fontName:(NSString *)fontName;
- (WordTextRange *) previousRangeBy:(WordE129)by count:(NSInteger)count;  // Returns a text range object relative to the specified text range.
- (WordTextRange *) previousSubdocument;  // Returns a new text range object to the previous subdocument. If there isn't another subdocument, an error occurs.
- (void) relocateDirection:(WordE224)direction;  // In outline view, moves the paragraphs within the text range after the next visible paragraph or before the previous visible paragraph. Body text moves with a heading only if the body text is collapsed in outline view or if it's part of the text range.
- (WordTextRange *) setRangeStart:(NSInteger)start end:(NSInteger)end;  // Returns a text range object by using the specified starting and ending character positions.
- (void) sortExcludeHeader:(BOOL)excludeHeader fieldNumber:(NSInteger)fieldNumber sortFieldType:(WordE178)sortFieldType sortOrder:(WordE179)sortOrder fieldNumberTwo:(NSInteger)fieldNumberTwo sortFieldTypeTwo:(WordE178)sortFieldTypeTwo sortOrderTwo:(WordE179)sortOrderTwo fieldNumberThree:(NSInteger)fieldNumberThree sortFieldTypeThree:(WordE178)sortFieldTypeThree sortOrderThree:(WordE179)sortOrderThree sortColumn:(BOOL)sortColumn separator:(WordE176)separator caseSensitive:(BOOL)caseSensitive languageID:(WordE182)languageID;  // Sorts the paragraphs in the specified text range.
- (void) sortAscending;  // Sorts paragraphs or table rows in ascending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (void) sortDescending;  // Sorts paragraphs or table rows in descending alphanumeric order. The first paragraph or table row is considered a header record and isn't included in the sort. Use the sort method to include the header record in a sort.
- (NSDictionary *) textRangeSpellingSuggestionsCustomDictionary:(WordDictionary *)customDictionary ignoreUppercase:(BOOL)ignoreUppercase mainDictionary:(WordDictionary *)mainDictionary suggestionMode:(WordE256)suggestionMode customDictionary2:(WordDictionary *)customDictionary2 customDictionary3:(WordDictionary *)customDictionary3 customDictionary4:(WordDictionary *)customDictionary4 customDictionary5:(WordDictionary *)customDictionary5 customDictionary6:(WordDictionary *)customDictionary6 customDictionary7:(WordDictionary *)customDictionary7 customDictionary8:(WordDictionary *)customDictionary8 customDictionary9:(WordDictionary *)customDictionary9 customDictionary10:(WordDictionary *)customDictionary10;  // Returns a record that contains the spelling error type and the list of words suggested as replacements for the first word in the specified range

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
@interface WordAutocorrectEntry : WordBaseObject

@property (copy) NSString *autocorrectValue;  // Returns or sets the value of the auto correct entry.
@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy) NSString *name;  // Returns or sets the name of the auto correct entry.
@property (readonly) BOOL richText;  // Returns if formatting is stored with the autocorrect entry replacement text.

- (void) applyCorrectionToRange:(WordTextRange *)toRange;  // Replaces a range with the value of the specified autocorrect entry.

@end

// Represents the autocorrect functionality in Word.
@interface WordAutocorrect : WordBaseObject

- (SBElementArray *) autocorrectEntries;
- (SBElementArray *) firstLetterExceptions;
- (SBElementArray *) twoInitialCapsExceptions;
- (SBElementArray *) otherCorrectionsExceptions;

@property BOOL correctDays;  // Returns or sets if Word automatically capitalizes the first letter of days of the week.
@property BOOL correctInitialCaps;  // Returns or sets if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, WOrd is corrected to Word.
@property BOOL correctSentenceCaps;  // Returns or sets if Word automatically capitalizes the first letter in each sentence.
@property BOOL correctTableCaps;  // Returns or sets if Word automatically capitalizes the first letter in each table cell.
@property BOOL firstLetterAutoAdd;  // Returns or sets if Word automatically adds abbreviations to the list of autocorrect first letter exceptions.
@property BOOL otherCorrectionsAutoAdd;  // Returns or sets if Microsoft Word automatically adds words to the list of other autocorrect exceptions. Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct.
@property BOOL replaceText;  // Returns or sets if Microsoft Word automatically replaces specified text with entries from the autocorrect list.
@property BOOL replaceTextFromSpellingChecker;  // Returns or sets if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types.
@property BOOL showAutocorrectSmartButton;  // Returns or sets if Word shows the AutoCorrect smart button which allows you to review the correction when an automatic correction occurs.
@property BOOL turnOnAutocorrect;  // Returns or sets if Word automatically corrects spelling and formatting as you type.
@property BOOL twoInitialCapsAutoAdd;  // Returns or sets if Microsoft Word automatically adds words to the list of autocorrect initial caps exceptions.


@end

// Represents a dictionary.
@interface WordDictionary : WordBaseObject

@property (readonly) WordE255 dictionaryType;  // Returns the dictionary type.
@property WordE182 languageID;  // Returns or sets the language for the dictionary object
@property BOOL languageSpecific;  // Returns or sets if the custom dictionary is to be used only with text formatted for a specific language.
@property (copy, readonly) NSString *name;  // Returns or sets the name of this dictionary object.
@property (copy, readonly) NSString *path;  // Returns the disk or Web path to the specified dictionary.
@property (readonly) BOOL readOnly;  // Returns true if the specified dictionary cannot be changed.


@end

// Represents an abbreviation excluded from automatic correction.
@interface WordFirstLetterException : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of this first letter exception object.


@end

// Represents a language used for proofing or formatting in Microsoft Word.
@interface WordLanguage : WordBaseObject

@property (copy, readonly) WordDictionary *activeGrammarDictionary;  // Returns a dictionary object that represents the active grammar dictionary for the specified language
@property (copy, readonly) WordDictionary *activeHyphenationDictionary;  // Returns a dictionary object that represents the active hyphenation dictionary for the specified language.
@property (copy, readonly) WordDictionary *activeSpellingDictionary;  // Returns a dictionary object that represents the active spelling dictionary for the specified language.
@property (copy, readonly) WordDictionary *activeThesaurusDictionary;  // Returns a dictionary object that represents the active thesaurus dictionary for the specified language.
@property (copy) NSString *defaultWritingStyle;  // Returns or sets the default writing style used by the grammar checker for the specified language. The name of the writing style is the localized name for the specified language.
@property (readonly) WordE182 languageID;  // Returns an enumeration that identifies the specified language.
@property (copy, readonly) NSString *name;  // Returns the name of the language
@property (copy, readonly) NSString *nameLocal;  // Returns the name of a proofing tool language in the language of the user.
@property WordE255 spellingDictionaryType;  // Returns or sets the proofing tool type
@property (copy, readonly) NSArray *writingStyleList;  // Returns a list of strings that contains the names of all writing styles available for the specified language.


@end

// Represents a single auto correct exception.
@interface WordOtherCorrectionsException : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // Returns the name of this other corrections exception object.


@end

// Represents one of the readability statistics for a document or range.
@interface WordReadabilityStatistic : WordBaseObject

@property (copy, readonly) NSString *name;  // The name of this readability statistic object.
@property (readonly) double readabilityValue;  // The value of this readability statistic object.


@end

// Represents the information about synonyms, antonyms, related words, or related expressions for the specified range or a given string.
@interface WordSynonymInfo : WordBaseObject

@property (copy, readonly) NSArray *antonyms;  // Returns a list of antonyms for the word or phrase.
@property (copy, readonly) NSString *context;  // Returns the word or phrase that was looked up by the thesaurus.
@property (readonly) BOOL found;  // Returns true if the thesaurus finds synonyms, antonyms, related words, or related expressions for the word or phrase.
@property (readonly) NSInteger meaningCount;  // Returns the number of entries in the list of meanings found in the thesaurus for the word or phrase. Returns zero if no meanings were found.
@property (copy, readonly) NSArray *meanings;  // Returns the list of meanings for the word or phrase.
@property (copy, readonly) NSArray *partOfSpeech;  // Returns a list of the parts of speech corresponding to the meanings found for the word or phrase looked up in the thesaurus.
@property (copy, readonly) NSArray *relatedExpressions;  // Returns a list of expressions related to the specified word or phrase. 
@property (copy, readonly) NSArray *relatedWords;  // Returns a list of words related to the specified word or phrase.

- (NSArray *) getSynonymListForItemToCheck:(NSString *)itemToCheck;  // Get the list of synonyms for a particular word
- (NSArray *) getSynonymListFromMeaningIndex:(NSInteger)meaningIndex;  // Get the list of synonyms using an index into the list of meanings

@end

// Represents a single initial-capital autocorrect exception.
@interface WordTwoInitialCapsException : WordBaseObject

@property (readonly) NSInteger entry_index;  // Returns the index for the position of the object in its container element list.
@property (copy, readonly) NSString *name;  // The name of this two initial caps exception object.


@end



/*
 * Table Suite
 */

// Represents a single table cell.
@interface WordCell : WordBaseObject

- (SBElementArray *) tables;

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordColumn *column;  // Returns the column object that contains this cell object.
@property (readonly) NSInteger columnIndex;  // Returns the number of the column that contains the specified cell.
@property BOOL fitText;  // Returns or sets if Microsoft Word visually reduces the size of text typed into a cell so that it fits within the column width.
@property double height;  // Returns or sets the height of the object.
@property WordE133 heightRule;  // Returns or sets the rule for determining the height of the specified cells.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified cell.
@property (copy, readonly) WordCell *nextCell;  // Returns the next cell object
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified cell.
@property WordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified column.
@property (copy, readonly) WordCell *previousCell;  // Returns the previous cell object
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordRow *row;  // Returns the row object that contains this cell object.
@property (readonly) NSInteger rowIndex;  // Returns the number of the row that contains the specified cell.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the cell object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the cell object.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property WordE147 verticalAlignment;  // Returns or sets the vertical alignment of text in one or more cells of a table.
@property double width;  // Returns or sets the width of the object.
@property BOOL wordWrap;  // Returns or set  if Microsoft Word wraps text to multiple lines and lengthens the cell so that the cell width remains the same.

- (void) autoSum;  // Inserts an = Formula field that calculates and displays the sum of the values in table cells above or to the left of the cell specified in the expression.
- (void) formulaFormulaString:(NSString *)formulaString numberFormatString:(NSString *)numberFormatString;  // Inserts an = Formula field that contains the specified formula into a table cell.
- (void) mergeCellWith:(WordCell *)with;  // Merges the specified table cell with another cell. The result is a single table cell.
- (void) splitCellNumberOfRows:(NSInteger)numberOfRows numberOfColumns:(NSInteger)numberOfColumns;  // Splits a single table cell into multiple cells.

@end

// Represents options that can be set for columns.
@interface WordColumnOptions : WordBaseObject

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double defaultWidth;  // Returns or sets the default width of columns.
@property (readonly) NSInteger nestingLevel;
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified columns. 
@property WordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified columns. 
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with columns.

- (void) distributeWidth;  // Adjusts the width of the specified columns or cells so that they're equal.

@end

// Represents a single table column.
@interface WordColumn : WordBaseObject

@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this column object.
@property (readonly) NSInteger columnIndex;  // Returns the index for the position of the object in its container element list.
@property (readonly) BOOL isFirst;  // Returns if the specified column is the first one in the table.
@property (readonly) BOOL isLast;  // Returns if the specified column is the last one in the table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified column.
@property (copy, readonly) WordColumn *nextColumn;  // Returns the next column object
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified column.
@property WordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified column.
@property (copy, readonly) WordColumn *previousColumn;  // Returns the previous column object
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the column object.
@property double width;  // Returns or sets the width of the object.


@end

@interface WordCondition : WordBaseObject

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
@interface WordRowOptions : WordBaseObject

@property WordE144 alignment;  // Returns or sets a constant that represents the alignment for rows.
@property BOOL allowBreakAcrossPages;  // Returns or sets if the text in a table row or rows are allowed to split across a page break.
@property BOOL allowOverlap;  // Returns or sets a value that specifies whether the specified rows can overlap other rows.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this cell object.
@property double distanceBottom;  // Returns or sets the distance in points between the document text and the bottom edge of the specified table. This property doesn't have any effect if wrap around text is false. 
@property double distanceLeft;  // Returns or sets the distance in points between the document text and the left edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property double distanceRight;  // Returns or sets the distance in points between the document text and the right edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property double distanceTop;  // Returns or sets the distance in points between the document text and the top edge of the specified table. This property doesn't have any effect if wrap around text is false.
@property BOOL headingFormat;  // Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page.
@property double height;  // Returns or sets the height of the object.
@property WordE133 heightRule;  // Returns or sets the rule for determining the height of the specified rows.
@property double horizontalPosition;  // Returns or sets the horizontal distance between the edge of the rows.
@property (readonly) NSInteger nestingLevel;
@property WordE236 relativeHorizontalPosition;  // Specifies to what the horizontal position of a group of rows is relative.
@property WordE237 relativeVerticalPosition;  // Specifies to what the vertical position of a group of rows is relative.
@property double rowLeftIndent;  // Returns or sets the left indent in points for the specified rows.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with rows.
@property double spaceBetweenColumns;  // Returns or sets the distance in points between text in adjacent columns of the specified row or rows.
@property double verticalPosition;  // Returns or sets the vertical distance between the edge of the rows.
@property BOOL wrapAroundText;  // Returns or sets whether text should wrap around the specified rows. 

- (void) distributeRowHeight;  // Adjusts the height of the specified rows or cells so that they're equal.

@end

// Represents a row in a table.
@interface WordRow : WordBaseObject

- (SBElementArray *) cells;

@property WordE144 alignment;  // Returns or sets a constant that represents the alignment for the specified row.
@property BOOL allowBreakAcrossPages;  // Returns or sets if the text in a row or rows are allowed to split across a page break.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property BOOL headingFormat;  // Returns or sets if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page.
@property double height;  // Returns or sets the height of the object.
@property WordE133 heightRule;  // Returns or sets the rule for determining the height of the specified rows.
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

@interface WordTableStyle : WordBaseObject

@property WordE144 alignment;  // Returns or sets a constant that represents the alignment for the specified row.
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
@interface WordTable : WordBaseObject

- (SBElementArray *) columns;
- (SBElementArray *) rows;
- (SBElementArray *) tables;

@property BOOL allowAutoFit;  // Returns or sets if Microsoft Word will automatically resize cells in a table to fit their contents.
@property BOOL allowPageBreaks;  // Returns or sets if Microsoft Word allows to break the specified table across pages.
@property BOOL applyStyleFirstColumn;
@property BOOL applyStyleHeadingRows;
@property BOOL applyStyleLastColumn;
@property BOOL applyStyleLastRow;
@property (readonly) WordE180 autoFormatType;  // Returns the type of automatic formatting that's been applied to the specified table.
@property (copy, readonly) WordBorderOptions *borderOptions;  // Returns back border options object associated with this table object.
@property double bottomPadding;  // Returns or sets the amount of space in points to add below the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordColumnOptions *columnOptions;  // Returns the column options object associated with this table object.
@property double leftPadding;  // Returns or sets the amount of space in points to add to the left of the contents of a single cell or all the cells in a table.
@property (readonly) NSInteger nestingLevel;  // Returns the nesting level of the specified table.
@property (readonly) NSInteger numberOfColumns;  // Returns the number of columns in this table
@property (readonly) NSInteger numberOfRows;  // Returns the number of rows in this table
@property double preferredWidth;  // Returns or sets the preferred width in points for the specified table.
@property WordE290 preferredWidthType;  // Returns or sets the preferred unit of measurement to use for the width of the specified table. 
@property double rightPadding;  // Returns or sets the amount of space in points to add to the right of the contents of a single cell or all the cells in a table.
@property (copy, readonly) WordRowOptions *rowOptions;  // Returns the row options object associated with this table object.
@property (copy, readonly) WordShading *shading;  // Returns the shading object associated with the table object.
@property double spacing;  // Returns or sets the spacing between the cells in a table.
@property WordE184 style;  // Returns or sets the Word style associated with the table object.
@property (copy, readonly) WordTextRange *textObject;  // Returns a text range object that represents the portion of a document that's contained in the table object.
@property double topPadding;  // Returns or sets the amount of space in points to add above the contents of a single cell or all the cells in a table.
@property (readonly) BOOL uniform;  // Returns if all the rows in a table have the same number of columns.

- (void) autoFitBehaviorBehavior:(WordE288)behavior;  // Determines how Microsoft Word resizes a table when the autofit feature is used. Word can resize the table based on the content of the table cells or the width of the document window.
- (void) autoFormatTableTableFormat:(WordE180)tableFormat applyBorders:(BOOL)applyBorders applyShading:(BOOL)applyShading applyFont:(BOOL)applyFont applyColor:(BOOL)applyColor applyHeadingRows:(BOOL)applyHeadingRows applyLastRow:(BOOL)applyLastRow applyFirstColumn:(BOOL)applyFirstColumn applyLastColumn:(BOOL)applyLastColumn autoFit:(BOOL)autoFit;  // Applies a predefined look to a table.
- (WordTextRange *) convertToTextSeparator:(WordE177)separator nestedTables:(BOOL)nestedTables;  // Converts a table to text and returns a text range object that represents the delimited text.
- (WordCell *) getCellFromTableRow:(NSInteger)row column:(NSInteger)column;  // Returns a cell object that represents a cell in a table.
- (WordTable *) splitTableRow:(NSInteger)row;  // Inserts an empty paragraph immediately above the specified row in the table, and returns a table object that contains both the specified row and the rows that follow it.
- (void) updateAutoFormat;  // Updates the table with the characteristics of a predefined table format. For example, if you apply a table format with auto format and then insert rows and columns, the table may no longer match the predefined look.

@end

