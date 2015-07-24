
declare enum TreeScope {
    Element = 0x1,
    Children = 0x2,
    Descendants = 0x4,
    Subtree = 0x7,
    Parent = 0x8,
    Ancestors = 0x10,
}

declare enum PatternId
{
    Invoke = 10000,
    Selection = 10001,
    Value = 10002,
    RangeValue = 10003,
    Scroll = 10004,
    ExpandCollapse = 10005,
    Grid = 10006,
    GridItem = 10007,
    MultipleView = 10008,
    Window = 10009,
    SelectionItem = 10010,
    Dock = 10011,
    Table = 10012,
    TableItem = 10013,
    Text = 10014,
    Toggle = 10015,
    Transform = 10016,
    ScrollItem = 10017,
    LegacyIAccessible = 10018,
    ItemContainer = 10019,
    VirtualizedItem = 10020,
    SynchronizedInput = 10021,
    ObjectModel = 10022,
    Annotation = 10023,
    TextPattern2 = 10024,
    Styles = 10025,
    Spreadsheet = 10026,
    SpreadsheetItem = 10027,
    TransformPattern2 = 10028,
    TextChild = 10029,
    Drag = 10030,
    DropTarget = 10031,
    TextEdit = 10032,
}

declare enum EventId
{
    ToolTipOpened = 20000,
    ToolTipClosed = 20001,
    StructureChanged = 20002,
    MenuOpened = 20003,
    AutomationPropertyChanged = 20004,
    AutomationFocusChanged = 20005,
    AsyncContentLoaded = 20006,
    MenuClosed = 20007,
    LayoutInvalidated = 20008,
    Invoke_Invoked = 20009,
    SelectionItem_ElementAddedToSelection = 20010,
    SelectionItem_ElementRemovedFromSelection = 20011,
    SelectionItem_ElementSelected = 20012,
    Selection_Invalidated = 20013,
    Text_TextSelectionChanged = 20014,
    Text_TextChanged = 20015,
    Window_WindowOpened = 20016,
    Window_WindowClosed = 20017,
    MenuModeStart = 20018,
    MenuModeEnd = 20019,
    InputReachedTarget = 20020,
    InputReachedOtherElement = 20021,
    InputDiscarded = 20022,
    SystemAlert = 20023,
    LiveRegionChanged = 20024,
    HostedFragmentRootsInvalidated = 20025,
    Drag_DragStart = 20026,
    Drag_DragCancel = 20027,
    Drag_DragComplete = 20028,
    DropTarget_DragEnter = 20029,
    DropTarget_DragLeave = 20030,
    DropTarget_Dropped = 20031,
    TextEdit_TextChanged = 20032,
    TextEdit_ConversionTargetChanged = 20033,
}

declare enum PropertyId
{
    RuntimeId = 30000,
    BoundingRectangle = 30001,
    ProcessId = 30002,
    ControlType = 30003,
    LocalizedControlType = 30004,
    Name = 30005,
    AcceleratorKey = 30006,
    AccessKey = 30007,
    HasKeyboardFocus = 30008,
    IsKeyboardFocusable = 30009,
    IsEnabled = 30010,
    AutomationId = 30011,
    ClassName = 30012,
    HelpText = 30013,
    ClickablePoint = 30014,
    Culture = 30015,
    IsControlElement = 30016,
    IsContentElement = 30017,
    LabeledBy = 30018,
    IsPassword = 30019,
    NativeWindowHandle = 30020,
    ItemType = 30021,
    IsOffscreen = 30022,
    Orientation = 30023,
    FrameworkId = 30024,
    IsRequiredForForm = 30025,
    ItemStatus = 30026,
    IsDockPatternAvailable = 30027,
    IsExpandCollapsePatternAvailable = 30028,
    IsGridItemPatternAvailable = 30029,
    IsGridPatternAvailable = 30030,
    IsInvokePatternAvailable = 30031,
    IsMultipleViewPatternAvailable = 30032,
    IsRangeValuePatternAvailable = 30033,
    IsScrollPatternAvailable = 30034,
    IsScrollItemPatternAvailable = 30035,
    IsSelectionItemPatternAvailable = 30036,
    IsSelectionPatternAvailable = 30037,
    IsTablePatternAvailable = 30038,
    IsTableItemPatternAvailable = 30039,
    IsTextPatternAvailable = 30040,
    IsTogglePatternAvailable = 30041,
    IsTransformPatternAvailable = 30042,
    IsValuePatternAvailable = 30043,
    IsWindowPatternAvailable = 30044,
    ValueValue = 30045,
    ValueIsReadOnly = 30046,
    RangeValueValue = 30047,
    RangeValueIsReadOnly = 30048,
    RangeValueMinimum = 30049,
    RangeValueMaximum = 30050,
    RangeValueLargeChange = 30051,
    RangeValueSmallChange = 30052,
    ScrollHorizontalScrollPercent = 30053,
    ScrollHorizontalViewSize = 30054,
    ScrollVerticalScrollPercent = 30055,
    ScrollVerticalViewSize = 30056,
    ScrollHorizontallyScrollable = 30057,
    ScrollVerticallyScrollable = 30058,
    SelectionSelection = 30059,
    SelectionCanSelectMultiple = 30060,
    SelectionIsSelectionRequired = 30061,
    GridRowCount = 30062,
    GridColumnCount = 30063,
    GridItemRow = 30064,
    GridItemColumn = 30065,
    GridItemRowSpan = 30066,
    GridItemColumnSpan = 30067,
    GridItemContainingGrid = 30068,
    DockDockPosition = 30069,
    ExpandCollapseExpandCollapseState = 30070,
    MultipleViewCurrentView = 30071,
    MultipleViewSupportedViews = 30072,
    WindowCanMaximize = 30073,
    WindowCanMinimize = 30074,
    WindowWindowVisualState = 30075,
    WindowWindowInteractionState = 30076,
    WindowIsModal = 30077,
    WindowIsTopmost = 30078,
    SelectionItemIsSelected = 30079,
    SelectionItemSelectionContainer = 30080,
    TableRowHeaders = 30081,
    TableColumnHeaders = 30082,
    TableRowOrColumnMajor = 30083,
    TableItemRowHeaderItems = 30084,
    TableItemColumnHeaderItems = 30085,
    ToggleToggleState = 30086,
    TransformCanMove = 30087,
    TransformCanResize = 30088,
    TransformCanRotate = 30089,
    IsLegacyIAccessiblePatternAvailable = 30090,
    LegacyIAccessibleChildId = 30091,
    LegacyIAccessibleName = 30092,
    LegacyIAccessibleValue = 30093,
    LegacyIAccessibleDescription = 30094,
    LegacyIAccessibleRole = 30095,
    LegacyIAccessibleState = 30096,
    LegacyIAccessibleHelp = 30097,
    LegacyIAccessibleKeyboardShortcut = 30098,
    LegacyIAccessibleSelection = 30099,
    LegacyIAccessibleDefaultAction = 30100,
    AriaRole = 30101,
    AriaProperties = 30102,
    IsDataValidForForm = 30103,
    ControllerFor = 30104,
    DescribedBy = 30105,
    FlowsTo = 30106,
    ProviderDescription = 30107,
    IsItemContainerPatternAvailable = 30108,
    IsVirtualizedItemPatternAvailable = 30109,
    IsSynchronizedInputPatternAvailable = 30110,
    OptimizeForVisualContent = 30111,
    IsObjectModelPatternAvailable = 30112,
    AnnotationAnnotationTypeId = 30113,
    AnnotationAnnotationTypeName = 30114,
    AnnotationAuthor = 30115,
    AnnotationDateTime = 30116,
    AnnotationTarget = 30117,
    IsAnnotationPatternAvailable = 30118,
    IsTextPattern2Available = 30119,
    StylesStyleId = 30120,
    StylesStyleName = 30121,
    StylesFillColor = 30122,
    StylesFillPatternStyle = 30123,
    StylesShape = 30124,
    StylesFillPatternColor = 30125,
    StylesExtendedProperties = 30126,
    IsStylesPatternAvailable = 30127,
    IsSpreadsheetPatternAvailable = 30128,
    SpreadsheetItemFormula = 30129,
    SpreadsheetItemAnnotationObjects = 30130,
    SpreadsheetItemAnnotationTypes = 30131,
    IsSpreadsheetItemPatternAvailable = 30132,
    Transform2CanZoom = 30133,
    IsTransformPattern2Available = 30134,
    LiveSetting = 30135,
    IsTextChildPatternAvailable = 30136,
    IsDragPatternAvailable = 30137,
    DragIsGrabbed = 30138,
    DragDropEffect = 30139,
    DragDropEffects = 30140,
    IsDropTargetPatternAvailable = 30141,
    DropTargetDropTargetEffect = 30142,
    DropTargetDropTargetEffects = 30143,
    DragGrabbedItems = 30144,
    Transform2ZoomLevel = 30145,
    Transform2ZoomMinimum = 30146,
    Transform2ZoomMaximum = 30147,
    FlowsFrom = 30148,
    IsTextEditPatternAvailable = 30149,
    IsPeripheral = 30150,
}

declare enum AttributeId
{
    AnimationStyle = 40000,
    BackgroundColor = 40001,
    BulletStyle = 40002,
    CapStyle = 40003,
    Culture = 40004,
    FontName = 40005,
    FontSize = 40006,
    FontWeight = 40007,
    ForegroundColor = 40008,
    HorizontalTextAlignment = 40009,
    IndentationFirstLine = 40010,
    IndentationLeading = 40011,
    IndentationTrailing = 40012,
    IsHidden = 40013,
    IsItalic = 40014,
    IsReadOnly = 40015,
    IsSubscript = 40016,
    IsSuperscript = 40017,
    MarginBottom = 40018,
    MarginLeading = 40019,
    MarginTop = 40020,
    MarginTrailing = 40021,
    OutlineStyles = 40022,
    OverlineColor = 40023,
    OverlineStyle = 40024,
    StrikethroughColor = 40025,
    StrikethroughStyle = 40026,
    Tabs = 40027,
    TextFlowDirections = 40028,
    UnderlineColor = 40029,
    UnderlineStyle = 40030,
    AnnotationTypes = 40031,
    AnnotationObjects = 40032,
    StyleName = 40033,
    StyleId = 40034,
    Link = 40035,
    IsActive = 40036,
    SelectionActiveEnd = 40037,
    CaretPosition = 40038,
    CaretBidiMode = 40039,
}

declare enum ControlTypeId
{
    ButtonControl = 50000,
    CalendarControl = 50001,
    CheckBoxControl = 50002,
    ComboBoxControl = 50003,
    EditControl = 50004,
    HyperlinkControl = 50005,
    ImageControl = 50006,
    ListItemControl = 50007,
    ListControl = 50008,
    MenuControl = 50009,
    MenuBarControl = 50010,
    MenuItemControl = 50011,
    ProgressBarControl = 50012,
    RadioButtonControl = 50013,
    ScrollBarControl = 50014,
    SliderControl = 50015,
    SpinnerControl = 50016,
    StatusBarControl = 50017,
    TabControl = 50018,
    TabItemControl = 50019,
    TextControl = 50020,
    ToolBarControl = 50021,
    ToolTipControl = 50022,
    TreeControl = 50023,
    TreeItemControl = 50024,
    CustomControl = 50025,
    GroupControl = 50026,
    ThumbControl = 50027,
    DataGridControl = 50028,
    DataItemControl = 50029,
    DocumentControl = 50030,
    SplitButtonControl = 50031,
    WindowControl = 50032,
    PaneControl = 50033,
    HeaderControl = 50034,
    HeaderItemControl = 50035,
    TableControl = 50036,
    TitleBarControl = 50037,
    SeparatorControl = 50038,
    SemanticZoomControl = 50039,
    AppBarControl = 50040,
}

declare enum AnnotationTypeId
{
    Unknown	= 60000,
    SpellingError = 60001,
    GrammarError = 60002,
    Comment = 60003,
    FormulaError = 60004,
    TrackChanges = 60005,
    Header = 60006,
    Footer = 60007,
    Highlighted = 60008,
}

declare enum StyleId
{
    Custom = 70000,
    Heading1 = 70001,
    Heading2 = 70002,
    Heading3 = 70003,
    Heading4 = 70004,
    Heading5 = 70005,
    Heading6 = 70006,
    Heading7 = 70007,
    Heading8 = 70008,
    Heading9 = 70009,
    Title = 70010,
    Subtitle = 70011,
    Normal = 70012,
    Emphasis = 70013,
    Quote = 70014,
    BulletedList = 70015,
    NumberedList = 70016,
}

declare enum ExpandCollapseId {
    Collapsed = 0,
    Expanded = 1,
    PartiallyExpanded = 2,
    LeafNode = 3
}

declare enum DockPositionId {
    Top = 0,
    Left = 1,
    Bottom = 2,
    Right = 3,
    Fill = 4,
    None = 5
}

declare enum ScrollAmount {
    LargeDecrement = 0,
    SmallDecrement = 1,
    NoAmount = 2,
    LargeIncrement = 3,
    SmallIncrement = 4
}

declare enum ToggleState {
    ToggleState_Off = 0,
    ToggleState_On = 1,
    ToggleState_Indeterminate = 2
}

declare enum SynchronizedInputType {
    KeyUp,
    KeyDown,
    LeftMouseUp,
    LeftMouseDown,
    RightMouseUp,
    RightMouseDown
}

declare enum RowOrColumnMajor {
    RowMajor,
    ColumnMajor,
    Indeterminate
}

declare enum TextPatternRangeEndpoint {
    Start,
    End
}

declare enum TextUnit {
   Character,
   Format,
   Word,
   Line,
   Paragraph,
   Page,
   Document
}

declare enum SupportedTextSelection {
    Multiple,
    None,
    Single
}

declare enum ZoomUnit {
    NoAmount = 0,
    LargeDecrement = 1,
    SmallDecrement = 2,
    LargeIncrement = 3,
    SmallIncrement = 4
}

declare enum WindowInteractionState {
    Running,
    Closing,
    ReadyForUserInteraction,
    BlockedByModalWindow,
    NotResponding
}

declare enum WindowVisualState {
    Normal,
    Maximized,
    Minimized
}

interface IUIAutomationElement {
    id: string;
    name: string;
    className: string;
    automationid: string;

    parentNode(): IUIAutomationElement;
    firstChild(): IUIAutomationElement;
    lastChild(): IUIAutomationElement;
    nextSibling(): IUIAutomationElement;
    previousSibling(): IUIAutomationElement;
    childNodes(): Array<IUIAutomationElement>;

    findFirst(condition: (el: IUIAutomationElement) => boolean, minDepth: number, maxDepth: number): IUIAutomationElement;
    findAll(condition: (el: IUIAutomationElement) => boolean, minDepth: number, maxDepth: number): Array<IUIAutomationElement>;

    setFocus();

    compare(el: IUIAutomationElement): boolean;

    getPattern(patternId: PatternId): any;
    getProperty(propertyId: PropertyId): any;

    addEventHandler(eventId: EventId, scope: TreeScope, handler: (el: IUIAutomationElement) => void);
    removeEventHandler(eventId: EventId, handler: (el: IUIAutomationElement) => void);

    addPropertyChangedEventHandler(scope: TreeScope, propertyArray: Array<number>[], handler: (el: IUIAutomationElement) => void);
    removePropertyChangedEventHandler(handler: (el: IUIAutomationElement) => void);

    addStructureChangedEventHandler(scope: TreeScope, handler: (el: IUIAutomationElement) => void);
    removeStructureChangedEventHandler(handler: (el: IUIAutomationElement) => void);
}

interface InvokePattern {
    invoke();
}

interface ValuePattern {
    isReadOnly() : boolean;
    getValue(): string;
    setValue(value: string);
}

interface ExpandCollapsePattern {
    getState(): ExpandCollapseId;
    expand();
    collapse();
}

interface DockPattern {
    getPosition(): DockPositionId;
    setPosition(val: DockPositionId);
}

interface SelectionPattern {
    canSelectMultiple(): boolean;
    isSelectionRequired(): boolean;
    getSelection(): Array<IUIAutomationElement>[];
}

interface RangeValuePattern {
    isReadOnly(): boolean;
    getValue(): number;
    settValue(val: number);
    largeChange(): number;
    smallChange(): number;
    maximum(): number;
    minimum(): number;
}

interface ScrollPattern {
    horizonatalScrollPercent(): number;
    verticalScrollPercent(): number;
    horizonatalViewSize(): number;
    verticalViewSize(): number;
    horizonatallyScrollable(): boolean;
    verticallyScrollable(): boolean;
    scroll(horizontalAmount: ScrollAmount, verticalAmount: ScrollAmount);
    setScrollPercent(horizontalPercent: number, verticalPercent: number);
}

interface TogglePattern {
    getState() : ToggleState;
    toggle()
}

interface AnnotationPattern {
    getType(): AnnotationTypeId;
    getName(): string;
    author(): string;
    dateTime(): string;
    target(): IUIAutomationElement;
}

interface DragPattern {
    isGrabbed(): boolean;
    getDropEffect(): string;
    getDropEffects(): Array<string>[];
    getGrabbedItems(): Array<IUIAutomationElement>[];
}

interface DropTargetPattern {
    getEffect(): string;
    getEffects(): Array<string>[];
}

interface GridItemPattern {
    getRow(): number;
    getColumn(): number;
    getRowSpan(): number;
    getColumnSpan(): number;
    getContainingGrid(): IUIAutomationElement;
}

interface GridPattern {
    getRowCount(): number;
    getColumnCount(): number;
    getItem(): IUIAutomationElement;
}

interface ItemContainerPattern {
    findItemByProperty(el: IUIAutomationElement, propertyId: PropertyId, val: any): IUIAutomationElement;
}

interface LegacyIAccessiblePattern {
    getChildId(): number;
    getDefaultAction(): string;
    getDescription(): string;
    getHelp(): string;
    getKeyboardShortcut(): string;
    getName(): string;
    getRole(): string;
    getSelection(): Array<IUIAutomationElement>[];
    getState(): number;
    doDefaultAction();
    select(flag:number);
    setValue(val: string);
}

interface MultipleViewPattern {
    getCurrentView(): number;
    setCurrentView(viewId: number);
    getSupportedView(): Array<number>[];
    getViewName(viewId: number): string;
}

interface ScrollItemPattern {
    scrollIntoView();
}

interface SelectionItemPattern {
    addToSelection();
    isSelected() : boolean;
    removeFromSelection();
    select();
    getSelectionContainer(): IUIAutomationElement;
}

interface SpreadsheetItemPattern {
    getFormula(): string;
    getAnnotationObjects(): Array<IUIAutomationElement>[];
    getAnnotationTypes(): Array<number>[];
}

interface SpreadsheetPattern {
    getItemByName(name: string): IUIAutomationElement;
}

interface StylesPattern {
    getExtendedProperties(): string;
    getFillColor(): number;
    getPatternColor(): number;
    getPatternStyle(): string;
    getShape(): string;
    getStyleId(): number;
    getStyleName(): string;
}

interface SynchronizedInputPattern {
    startListening(inputType: SynchronizedInputType);
    cancel();
}

interface TableItemPattern {
    getColumnHeaderItems(): Array<IUIAutomationElement>[];
    getRowHeaderItems(): Array<IUIAutomationElement>[];
}

interface TablePattern {
    getRowOrColumnMajor(): RowOrColumnMajor;
    getColumnHeaders(): Array<IUIAutomationElement>[];
    getRowHeaders(): Array<IUIAutomationElement>[];
}

interface VirtualizedItemPattern {
    realize();
}

interface TextRange
{
    addToSelection();
    clone(): TextRange;
    compare(range: TextRange) : boolean;
    compareEndPoints(endpoint: TextPatternRangeEndpoint, target: TextRange, targetEndpoint: TextPatternRangeEndpoint): number;
    expandToEncloseUnit(unit: TextUnit);
    findAttribute(attributeId: AttributeId, val: any, backward: boolean): TextRange;
    findText(text: string, backward: boolean, ignoreCase: boolean): TextRange;
    getAttributeValue(attributeId: AttributeId): any;
    getBoundingRectangles(): Array<number>[];
    getChildred(): Array<IUIAutomationElement>[];
    getEnclosingElement(): IUIAutomationElement;
    getText(): string;
    move(unit: TextUnit, count: number): number;
    moveEnpointByRange(endpoint: TextPatternRangeEndpoint, target: TextRange, targetEndpoint: TextPatternRangeEndpoint);
    moveEndpointByUnit(endpoint: TextPatternRangeEndpoint, unit: TextUnit, count: number): number;
    removeFromSelection();
    scrollIntoView(alignToTop: boolean);
    select();

    showContextMenu();
}

interface TextChildPattern {
    getTextContainer(): IUIAutomationElement;
    getTextRange(): TextRange;
}

interface TextEditPattern {
    getActiveComposition(): TextRange;
    getConversationTarget(): TextRange;
}

interface TextPattern {
    getDocumentRange(): TextRange;
    getSupportedTextSelection(): SupportedTextSelection;
    getSelection(): Array<TextRange>[];
    getVisibleRange(): Array<TextRange>[];
    getRangeFromChild(child: IUIAutomationElement): TextRange;
    getRangeFromPoint(x: number, y: number): TextRange;

    getCaretRange(isActive: boolean): TextRange;
    getRangeFromAnnotation(annotationElement: IUIAutomationElement): TextRange;
}

interface TransformPattern {
    getCanMove(): boolean;
    getCanResize(): boolean;
    getCanRotate(): boolean;
    move(x: number, y: number);
    resize(x: number, y: number);
    rotate(degrees: number);

    getCanZoom(): boolean;
    zoom(zoom: number);
    zoomByUnit(zoom: ZoomUnit);
    getZoomLevel(): number;
    getZoomMaximum(): number;
    getZoomMinimum(): number;
}

interface WindowPattern
{
    getWindowInteractionState(): WindowInteractionState;
    isModal(): boolean;
    isTopmost(): boolean;
    canMaximize(): boolean;
    canMinimize(): boolean;
    getWindowVisualState(): WindowVisualState;
    setWindowVisualState(state: WindowVisualState);
    waitForInputIdle(ms: number): boolean;
}

interface IUIAutomation {
    root(): IUIAutomationElement;
    focused(): IUIAutomationElement;
    fromPoint(x: number, y: number): IUIAutomationElement;
}

interface Point {
    x: number;
    y: number;
}

interface Event {
    altKey: boolean;
    ctrlKey: boolean;
    shiftKey: boolean;

    keyCode: number;

    buttons: number;
    pos: Point;
}

interface Console {
    log(message: string);
}

interface Narrator {
    say(msg: string);
    sendInput(msg: string);
    stopReading();
    doPrimaryAction();
    moveToNextItem();
    moveToPreviousItem();
    changeVerbosityMode();
    showCommandsList();
    showCommandsForCurrentItem();
    jumpToNextCellInRow();
    jumpToPreviousCellInRow();
    jumpToNextCellInColumn();
    jumpToPreviousCellInColumn();
    getRow();
    getColumn();
    jumpToTableCell();
    getColumnHeader();
    getRowHeader();
    toggleTouchMode();
    toggleKeystrokeAnnouncements();
    lockNarratorKey();
    passKeysToApp();
    repeatLastPhrase();
    increaseVoiceVolume();
    decreaseVoiceVolume();
    increaseVoiceSpeed();
    decreaseVoiceSpeed();
    readItem();
    readItemAdvanced();
    readItemSpelled();
    readWindow();
    readAllItemsInContainingArea();
    moveToLastItemInContainingArea();
    moveNarratorCursorToSystemCursor();
    moveNarratorCursorToPointer();
    setFocusToItem();
    goBackOneItem();
    jumpToLinkedItem();
    startReading();
    readTextFromStartToCursor();
    readTextAttributes();
    readDocument();
    readNextPage();
    readCurrentPage();
    readPreviousPage();
    readNextParagraph();
    readCurrentParagraph();
    readPreviousParagraph();
    readNnextLine();
    readCurrentLine();
    readPreviousLine();
    readNextWord();
    readCurrentWord();
    readPreviousWord();
    readNextCharacter();
    readCurrentCharacter();
    readPreviousCharacter();
    jumpToNextHeading();
    jumpToPreviousHeading();
    jumpToNextTable();
    jumpToPreviousTable();
    jumpToNextLink();
    jumpToPreviousLink();
    moveToBeginningOfText();
    moveToEndOfText();
    readCurrentDateAndTime();
}

interface Host {
    onKeypress: (e: Event) => void;
    onActivate: () => void;
    onDeactivate: () => void;
    onSetFocus: () => void;
    onKillFocus: () => void;

    Storage: any;
}

declare var uia: IUIAutomation;
declare var host: Host;
declare var console: Console;
declare var narrator: Narrator;

