VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHyperList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***************************************************************************************
'*  HyperList!   Virtual Listview and Storage Container 3.0                            *
'*                                                                                     *
'*  Created:     June 23, 2006                                                         *
'*  Updated:     June 29, 2006                                                         *
'*  Purpose:     Ultra-Fast Virtual Listview and Storage Manipulation Class            *
'*  Functions:   (listed)                                                              *
'*  Revision:    3.0.0                                                                 *
'*  Compile:     Native                                                                *
'*  Referenced:  mData                                                                 *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************



' ~*** Exposed Functions ***~

'/~ CreateList              - initialize the listview           [in -long (5) | out -bool]
'/~ ColumnAdd               - create column headers             [in -long|string (5) | out -bool]
'/~ ColumnRemove            - remove a column                   [in -long | out -bool]
'/~ ColumnAutosize          - autosize columns                  [in -long (2) | out -bool]
'/~ InitImlHeader           - initialize header imagelist       [in -none | out -bool]
'/~ ImlHeaderAddBmp         - add a bitmap to header iml        [in -long (2) | out -bool]
'/~ ImlHeaderAddIcon        - add an icon to header iml         [in -long | out -long]
'/~ DestroyImlHeader        - destroy header image list         [in -none | out -bool]
'/~ InitImlSmall            - initialize smallicons image list  [in -none | out -col]
'/~ ImlSmallAddBmp          - add bmp to small image iml        [in -long (2) | out -bool]
'/~ ImlSmallAddIcon         - add icon to small image iml       [in -long | out -long]
'/~ DestroyImlSmall         - destroy small icons image list    [in -none | out -bool]
'/~ InitImlLarge            - initialize large icons image list [in -long (2) | out -bool]
'/~ ImlLargeAddBmp          - add bmp to large image iml        [in -long (2) | out -long]
'/~ ImlLargeAddIcon         - add icon to large image iml       [in -long | out -long]
'/~ DestroyImlLarge         - destroy large icons image list    [in -none | out -bool]
'/~ ItemAdd                 - add a single item to the list     [in -long (3)| string (1) | out -bool]
'/~ ItemRemove              - remove an item from the list      [in -long | out -bool]
'/~ ItemEnsureVisible       - move to item index                [in -long | out -bool]
'/~ LoadArray               - load data structure               [in -none | out -bool]
'/~ CreateList              - load data structure               [in -none | out -bool]
'/~ ColumnClear             - remove all columns                [in -none | out -bool]
'/~ ItemsSort               - sort items                        [in -long | out -bool]


' ~*** Exposed Properties ***~

'/~ p_StructPntr            - get/set pointer to the data structure
'/~ p_ColumnText            - get/set a columns heading
'/~ p_ColumnWidth           - get/set a columns length
'/~ p_ColumnAlign           - get/set a columns text alignment
'/~ p_ColumnIcon            - get/set header icon index
'/~ p_ColumnCount           - get/set retieve column count
'/~ HeaderHwnd              - return the column header handle
'/~ p_HeaderColor           - get/set return the header color
'/~ p_HeaderForeColor       - get/set return the header forecolor
'/~ p_HeaderCustom          - return the custom header status
'/~ p_Count                 - get item count
'/~ p_ItemsSorted           - get/set sorted mode status
'/~ p_IndirectMode          - get/set indirect mode status
'/~ p_ItemText              - get/set item text
'/~ p_ItemIcon              - get/set return icon index
'/~ p_ItemIndent            - get/set item indent
'/~ p_ItemSelected          - get selected state
'/~ p_ItemFocused           - get item focused state
'/~ p_ItemChecked           - get item checked state
'/~ p_ItemGhosted           - get/set item ghosted state
'/~ p_SubItemSet            - get/set a subitem
'/~ p_SubItemText           - get/set subitem text
'/~ p_SubItemIcon           - get/set subitem icon
'/~ p_SubItemImages         - get/set subitem icon state
'/~ p_BackColor             - get/set list backcolor
'/~ p_BorderStyle           - get/set list borderstyle
'/~ p_CheckBoxes            - get/set checkbox state
'/~ p_Font                  - get/set list font
'/~ p_ForeColor             - get/set list forecolor
'/~ p_FullRowSelect         - get/set full row select state
'/~ p_GridLines             - get/set gridlines state
'/~ p_HeaderDragDrop        - get/set drag and drop state
'/~ p_HeaderFixedWidth      - get/set fixed width state
'/~ p_HeaderFlat            - get/set header flat state
'/~ p_HeaderHide            - get/set header visible state
'/~ p_HideSelection         - get/set selection visible state
'/~ p_LabelTips             - get/set label tips state
'/~ p_MultiSelect           - get/set multiselect state
'/~ p_OneClickActivate      - get/set oneclick state
'/~ p_ScrollBarFlat         - get/set scrollbar state
'/~ p_SelectedCount         - get selected count
'/~ p_ViewMode              - get/set list viewmode state



'~*** Notes ***~

'* Credits/Cudos/More info..
'~ A big shout out to Carles 'da man!' PV, for his awesome api listview:
'~ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56021&lngWId=1
'~ Much of the HyperList class was derived from Carles example, and without Carles great demo of styles and methods
'~ this demonstration would have been much harder, (if not impossible), to create.
'~ Steve 'big Steve' McMahon, and his unparalleled listview control:
'~ http://www.vbaccelerator.com/home/VB/Code/Controls/ListView/article.asp
'~ if it is possible to do it with a listview, Steve has demonstrated it with this control, and as always,
'~ a great wealth of information and inspiration lies in the source code of this control.
'~ Rohan 'the Sort Monster' RDE, for his incredible QSort routines:
'~ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63800&lngWId=1
'~ The Qsort routines are a simply inspirational use of refined logic and methods.
'~ and of course, as per usual, a (weak and confusing) reference per M$:
'~ http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/listview_overview.asp?frame=true

'* Where it began..
'~ I first realized I had a need for something like this while benching an application I am currently writing.
'~ The app, could return 20k results in .2 seconds, ('that's a good thing' tm.), but, it took 9 seconds to write results
'~ into a highly optimized grid control, (that is a bad thing ;o()
'~ This is when I began my search for a faster method of committing results to a data container control.
'~ I had heard of a virtual list, it is something used primarily in the C++ world for linking large databases to a list control.
'~ I searched the internet for hours, but the only example of this I could find was a rather dated VB5 example on
'~ mvps.org. This example was incomplete, there was no storage solution, no sorting, a very old, (and very unstable), subclassing component,
'~ and most of the listview properties simply weren't implemented. The code was also, (I think unnecessarily..), complex,
'~ perhaps the reason why on one has expanded on it for some time. The mvps example uses a listview embedded in a user control
'~ to manage callback information, and an object pointer 'swap' (that still has me confused.. ;o).
'~ I decided to do this in a completely different way, the way you would create a listview in C++, by using CreateWindowEx
'~ to invoke the listview class, and adding the LVS_OWNERDATA flag, which lets the list know, data will be managed externally.
'~ Most of the other methods used are simply passing flags into the list via sendmessage, with a few notable exceptions..
'~ Checkboxes are a serious pain to implement in a virtual listview, simply passing the flag into a list does nothing, rather
'~ it is necessary to track the item state, and intercept callbacks to parse the checkbox click event, then simply mark the
'~ check state in an array, and then invalidate the list, forcing a repaint.
'~ Quite obviously, sorting and storage are also handled very differently. Very little of the listviews internal
'~ structures are used, data is not copied into the list, and can not be manipulated through the lists interfaces.
'~ So methods had to be constructed to manage these functions. The Item and SubItem arrays can be used to
'~ sort the data, with sorted data then being accessed through an indexed pointer array.
'~ These methods are far faster then using a listviews built in methods, (ever try to sort 100,000 items in a listview?
'~ wouldn't recommend it.. ;o)
'~ The memory access type lib, includes various functions to both memory and string functions, and the internal vb
'~ PutMem/GetMem functions. Benches demonstrated that on a compiled library, the type lib was consistantly faster
'~ then using api declarations. (experiment with caution :~0}

'* What can it do?
'~ Well.. it is simply the fastest listview control it is possible to make! It can load a million items in
'~ the blink of an eye, or sort 100 thousand items in a fraction of a second.
'~ This is, (after many hours of searching the net would seem to indicate..), the only fully implemented
'~ example of a virtual listview in vb. It has a complete sorting and storage apparatus, that are about as
'~ fast as you can make them in this language.
'~ It is ownerdrawn, which means that styles can be manipulated far beyond what can be accomplished with
'~ a standard listview.
'~ Headers are subclassed as a simple demo, but could be expanded dramatically, (like tiling image in
'~ place of header for modern mp3 player styles). Sky is the limit people, where you take it, is up to you..

'* Why a Dll, and not an Ocx?
'~ One thing I have noticed about people who write in this language, (that I have always found to
'~ be a bit puzzling..), is a singular aversion to using libraries, be it type libs or active-x dlls.
'~ I find it strange because.. the entire windows platform(s) are based around Dll/Class communication!
'~ Every so often I see a comment or posting about the superiority of creating a software device with no
'~ dependencies, and it seems that quite a few people believe that this is somehow true. If, you can
'~ build the component in this way, without sacrificing a single dolop of speed, flexibility or power,
'~ then this is indeed the best method. However, if you can make your program run significantly faster by
'~ using a library, type lib, xls file, whatever.. but will not, because you are resistant to adding
'~ external dependencies, then your design has failed by not producing the best software possible.
'~ Some will argue that adding dependencies creates new points of vulnerability, and complicates
'~ installation scripts, but really, I think it is more motivated by a mental disconnect that seems unique
'~ to the vb world. C++ designers commonly use support libraries without a second thought, their choice only
'~ motivated by what will work better, and what is the more efficient method.
'~ Users do not care about file dependencies, (they most likely have no idea what a .tlb or .dll file is..).
'~ If you look at any complex piece of software, say Photoshop or MS office, you will see dozens of
'~ support libraries, again, designers were only concerned with most efficient software they could
'~ produce, and not bound by ridiculous and irrelevant hangups. So my advice is: get over it,
'~ and let performance be the final arbiter in these matters.
'~ Another thing that causes me to roll my eyes.. ;o), is this notion of writing in 'pure vb'. -This
'~ is absolute, unadolterated rubbish! There is no such thing as 'pure vb' my friends, it is yet another
'~ silly notion people have, that can have a serious impact on the quality of software they produce.
'~ For instance, many of the common tasks you -think- are being performed by vb, are not being done
'~ by vb at all, but vb is only acting as a wrapper to system libraries, (ex. reg and file access).
'~ So even if you dedicate yourself to vb methods only, and completely forego api, you are still
'~ using system libraries, only with an additional layer of indirection.
'~ Here is a good example of when/why not to use vb methods..
'~ a) Vb Method:    String1 = String2 <- vb method, conversion/copying makes this very, very slow..
'~ b) CopyMemory:   CopyMemory StrPtr(String1), StrPtr(String2), &H4 <- upto 5x+ faster then method a)
'~ c) TypeLib:      CopyMemBv StrPtr(String1), StrPtr(String2), &H4  <- Fastest! upto 10% faster then b) when compiled
'~ Benched these in loops of 10k iterations using queryperf.
'~ So why?! Why would you use vb method?! It makes no sense!! You can do it many times faster using api or typelib,
'~ does it not make sense to use the more efficient method? What possible reason could there be to use vb method,
'~ when other methods will produce far better results? Emotional attatchment to vb? (har har.. ;o)

'* The only real advantage to using an ocx in this demonstration, would have been a property sheet.
'~ This would allow the user to set some default settings with a mouse click, rather then programmatically.
'~ However, advantages of using a dll are many.. A dll typically processes faster then an exe or ocx, has less
'~ overhead, and a smaller compile size, (and with a little work, could be made asynchronous).

'* The Truth about Type Libraries, (getting the most out of vb..)
'~ Nothing in vb, (or programming for that matter..), is done automagically. There is always some bit of
'~ programming logic, somewhere, somehow, managing and interpreting any action being performed, regardless
'~ of how trivial it may appear outwardly. That being said, lets look at the way vb handles api and declarations
'~ within a module.
'~ Api function pointers are resolved at runtime, (even on a compiled exe), this means that when a module loads,
'~ reference tables are generated that contain pointers to each and every api call in the module.
'~ When an api function is called, a lookup is performed in the function table to resolve the functions pointer
'~ address. This lookup costs you in time, (and the more declarations, the larger the table, and longer the lookup
'~ time, - and the worst thing possible, is to put _all_ your api functions in a module for global access, then
'~ you are adding another layer of indirection by looking up the owners module address, and searching through a
'~ sometimes very large global table of function pointers.
'~ A type library does not have this issue. Function pointers are resolved at compile time to the routine,
'~ (-no lookup is required-).
'~ One of the biggest limitations of vb is its inability to manage some data types, (like ex. IUnknown). Because a
'~ type lib is language independant, it can access any native data type, so api types that require data type arguments
'~ that are unknown to the runtime, can be accessed through a type interface.
'~ Part of the vb api management is a getlasterror check each time an api call is made, so this adds
'~ a second, (and often unneeded) call to every api function, but.. you bypass this entirely when using a type library.
'~ Type libs can also bypass expensive operations done automatically by the runtime when a function call is
'~ made, a good example of this is the implict ansi/unicode conversion done when vb passes a string into a function.
'~ Type libs will help to compile to a smaller file. The linker does not include unused functions, (unlike declare statements).

'~ Questions, (let's get them out of the way now..)
'~ Could this be made into a usercontrol?
'~ Yes. It would take little more then dropping this class into a uc, and changing callback pointer to the
'~ uc hwnd. But, you will lose some speed and performance, (have you not heard anything I've said?!).
'~ Could this be made without the type library dependencies?
'~ Perhaps, but again, you are bartering a lot of speed and power for a very small convenience only.
'~ Could this be linked directly to a database?
'~ Yes. Should be simple enough to create a connector to an indexed database for someone who knows
'~ what they are doing. I have left the eHIndirect callback event as an option for this purpose.

'~ Release 1.0 - June 23, 2006
'~ Still a few todo's.. icons aren't working right yet, not sure why..
'~ A couple other routines need some, work, maybe this weekend..

'~ Release 2.0 - June 26, 2006
'~ Icon issues (partially) resolved, a few fixes here and there, and a uc version for people who can't live without it..

'~ Release 3.0 - June 29, 2006
'~ Icon issues are resolved, (in both versions). Added tab focus intercept in dll version, and ipao in
'~ uc version. Leak in dll harness resolved. Unless bugs are found, or I add a couple more features down
'~ the road, this should be it.. (3000 times faster then a standard listview, not bad.. ;o)

'~ Enjoy..
'~ Comment or a job: steppenwolfe_2000@yahoo.com


Implements MISubclass
'/* note: I would remove every unused constant/declaration in finished library
'/* I left the bulk of them in for a reference

Private Const NEG1                                  As Long = -1
Private Const m0                                    As Long = &H0
Private Const m1                                    As Long = &H1
Private Const m2                                    As Long = &H2
Private Const m4                                    As Long = &H4
Private Const m8                                    As Long = &H8
Private Const m32                                   As Long = &H20

Private Const API_TRUE                              As Long = 1&
Private Const CLR_NONE                              As Long = &HFFFFFFFF

Private Const H_MAX                                 As Long = &HFFFF + 1
Private Const NM_FIRST                              As Long = H_MAX
Private Const NM_CUSTOMDRAW                         As Long = (NM_FIRST - 12)
Private Const CDDS_PREPAINT                         As Long = &H1
Private Const CDDS_POSTPAINT                        As Long = &H2
Private Const CDDS_PREERASE                         As Long = &H3
Private Const CDDS_POSTERASE                        As Long = &H4
Private Const CDDS_ITEM                             As Long = &H10000
Private Const CDDS_ITEMPREPAINT                     As Long = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT                    As Long = CDDS_ITEM Or CDDS_POSTPAINT
Private Const CDDS_ITEMPREERASE                     As Long = CDDS_ITEM Or CDDS_PREERASE
Private Const CDDS_ITEMPOSTERASE                    As Long = CDDS_ITEM Or CDDS_POSTERASE
Private Const CDDS_SUBITEM                          As Long = &H20000

Private Const CDRF_DODEFAULT                        As Long = &H0
Private Const CDRF_NEWFONT                          As Long = &H2
Private Const CDRF_SKIPDEFAULT                      As Long = &H4
Private Const CDRF_NOTIFYPOSTPAINT                  As Long = &H10
Private Const CDRF_NOTIFYITEMDRAW                   As Long = &H20
Private Const CDRF_NOTIFYSUBITEMDRAW                As Long = &H20
Private Const CDRF_NOTIFYPOSTERASE                  As Long = &H40
Private Const CDRF_NOTIFYITEMERASE                  As Long = &H80

Private Const FW_NORMAL                             As Long = 400
Private Const FW_BOLD                               As Long = 700

Private Const GWL_STYLE                             As Long = (-16)
Private Const GWL_EXSTYLE                           As Long = (-20)

Private Const HDF_LEFT                              As Long = 0
Private Const HDF_RIGHT                             As Long = 1
Private Const HDF_CENTER                            As Long = 2
Private Const HDF_IMAGE                             As Long = &H800
Private Const HDF_STRING                            As Long = &H4000
Private Const HDF_BITMAP_ON_RIGHT                   As Long = &H1000

Private Const HDI_WIDTH                             As Long = &H1
Private Const HDI_TEXT                              As Long = &H2
Private Const HDI_FORMAT                            As Long = &H4
Private Const HDI_IMAGE                             As Long = &H20

Private Const HDN_FIRST                             As Long = -300
Private Const HDN_ITEMCHANGING                      As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMCLICK                         As Long = (HDN_FIRST - 2)
Private Const HDN_ITEMDBLCLICK                      As Long = (HDN_FIRST - 3)
Private Const HDN_DIVIDERDBLCLICK                   As Long = (HDN_FIRST - 5)
Private Const HDN_BEGINTRACK                        As Long = (HDN_FIRST - 6)
Private Const HDN_ENDTRACK                          As Long = (HDN_FIRST - 7)
Private Const HDN_TRACK                             As Long = (HDN_FIRST - 8)
Private Const HDN_GETDISPINFO                       As Long = (HDN_FIRST - 9)
Private Const HDN_BEGINDRAG                         As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG                           As Long = (HDN_FIRST - 11)

Private Const HDM_FIRST                             As Long = &H1200
Private Const HDM_GETITEMCOUNT                      As Long = (HDM_FIRST + 0)
Private Const HDM_INSERTITEM                        As Long = (HDM_FIRST + 1)
Private Const HDM_DELETEITEM                        As Long = (HDM_FIRST + 2)
Private Const HDM_GETITEM                           As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEM                           As Long = (HDM_FIRST + 4)
Private Const HDM_SETIMAGELIST                      As Long = (HDM_FIRST + 8)

Private Const HDS_BUTTONS                           As Long = &H2

Private Const ILC_COLOR                             As Long = &H0
Private Const ILC_MASK                              As Long = &H1
Private Const ILC_COLOR4                            As Long = &H4
Private Const ILC_COLOR8                            As Long = &H8
Private Const ILC_COLOR16                           As Long = &H10
Private Const ILC_COLOR24                           As Long = &H18
Private Const ILC_COLOR32                           As Long = &H20

Private Const LOGPIXELSY                            As Long = 90

Private Const LVCF_FMT                              As Long = &H1
Private Const LVCF_WIDTH                            As Long = &H2
Private Const LVCF_TEXT                             As Long = &H4
Private Const LVCF_SUBITEM                          As Long = &H8
Private Const LVCF_IMAGE                            As Long = &H10
Private Const LVCF_ORDER                            As Long = &H20

Private Const LVM_FIRST                             As Long = &H1000
Private Const LVM_GETBKCOLOR                        As Long = (LVM_FIRST + 0)
Private Const LVM_SETBKCOLOR                        As Long = (LVM_FIRST + 1)
Private Const LVM_GETIMAGELIST                      As Long = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST                      As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT                      As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEM                           As Long = (LVM_FIRST + 5)
Private Const LVM_SETITEM                           As Long = (LVM_FIRST + 6)
Private Const LVM_INSERTITEM                        As Long = (LVM_FIRST + 7)
Private Const LVM_DELETEITEM                        As Long = (LVM_FIRST + 8)
'Private Const LVM_DELETEALLITEMS                    As Long = (LVM_FIRST + 9)
Private Const LVM_GETNEXTITEM                       As Long = (LVM_FIRST + 12)
Private Const LVM_FINDITEM                          As Long = (LVM_FIRST + 13)
'Private Const LVM_HITTEST                           As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE                     As Long = (LVM_FIRST + 19)
'Private Const LVM_SCROLL                            As Long = (LVM_FIRST + 20)
Private Const LVM_REDRAWITEMS                       As Long = (LVM_FIRST + 21)
Private Const LVM_ARRANGE                           As Long = (LVM_FIRST + 22)
Private Const LVM_EDITLABEL                         As Long = (LVM_FIRST + 23)
Private Const LVM_GETEDITCONTROL                    As Long = (LVM_FIRST + 24)
Private Const LVM_GETCOLUMN                         As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMN                         As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMN                      As Long = (LVM_FIRST + 27)
Private Const LVM_DELETECOLUMN                      As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH                    As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH                    As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                         As Long = (LVM_FIRST + 31)
Private Const LVM_GETTEXTCOLOR                      As Long = (LVM_FIRST + 35)
Private Const LVM_SETTEXTCOLOR                      As Long = (LVM_FIRST + 36)
Private Const LVM_GETTEXTBKCOLOR                    As Long = (LVM_FIRST + 37)
Private Const LVM_SETTEXTBKCOLOR                    As Long = (LVM_FIRST + 38)
Private Const LVM_UPDATE                            As Long = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE                      As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE                      As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT                       As Long = (LVM_FIRST + 45)
Private Const LVM_SETITEMTEXT                       As Long = (LVM_FIRST + 46)
Private Const LVM_SORTITEMS                         As Long = (LVM_FIRST + 48)
Private Const LVM_GETSELECTEDCOUNT                  As Long = (LVM_FIRST + 50)
'Private Const LVM_SETEXTENDEDLISTVIEWSTYLE          As Long = (LVM_FIRST + 54)
'Private Const LVM_GETEXTENDEDLISTVIEWSTYLE          As Long = (LVM_FIRST + 55)
Private Const LVM_SETHOTITEM                        As Long = (LVM_FIRST + 60)
Private Const LVM_GETHOTITEM                        As Long = (LVM_FIRST + 61)
Private Const LVM_SETHOTCURSOR                      As Long = (LVM_FIRST + 62)
Private Const LVM_GETHOTCURSOR                      As Long = (LVM_FIRST + 63)
Private Const LVM_SETBKIMAGE                        As Long = (LVM_FIRST + 68)
Private Const LVM_GETBKIMAGE                        As Long = (LVM_FIRST + 69)
Private Const LVM_SETVIEW                           As Long = (LVM_FIRST + 142)
Private Const LVM_GETVIEW                           As Long = (LVM_FIRST + 143)

'Private Const LVIF_TEXT                            As Long = &H1
'Private Const LVIF_IMAGE                           As Long = &H2
'Private Const LVIF_PARAM                           As Long = &H4
'Private Const LVIF_STATE                           As Long = &H8
'Private Const LVIF_INDENT                          As Long = &H10
Private Const LVIF_GROUPID                          As Long = &H100
Private Const LVIF_COLUMNS                          As Long = &H200

'Private Const LVIS_STATEIMAGEMASK                  As Long = &HF000
'Private Const LVIS_FOCUSED                         As Long = &H1
'Private Const LVIS_SELECTED                        As Long = &H2
'Private Const LVIS_CUT                             As Long = &H4
'Private Const LVIS_DROPHILITED                     As Long = &H8
'Private Const LVIS_OVERLAYMASK                     As Long = &HF00

Private Const LVCFMT_LEFT                           As Long = &H0
Private Const LVCFMT_RIGHT                          As Long = &H1
Private Const LVCFMT_CENTER                         As Long = &H2
Private Const LVCFMT_JUSTIFYMASK                    As Long = &H3
Private Const LVCFMT_IMAGE                          As Long = &H800
Private Const LVCFMT_BITMAP_ON_RIGHT                As Long = &H1000
Private Const LVCFMT_COL_HAS_IMAGES                 As Long = &H8000
'Private Const  INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
Private Const LVIS_CHECKED                          As Long = &H2000&
Private Const LVIS_UNCHECKED                        As Long = &H1000&
Private Const LVIS_CHKCLICK                         As Long = &HFFFE
Private Const LVS_ICON                              As Long = &H0
Private Const LVS_REPORT                            As Long = &H1
Private Const LVS_SMALLICON                         As Long = &H2
Private Const LVS_LIST                              As Long = &H3

Private Const LVS_EX_GRIDLINES                      As Long = &H1&
Private Const LVS_EX_SUBITEMIMAGES                  As Long = &H2&
Private Const LVS_EX_CHECKBOXES                     As Long = &H4&
Private Const LVS_EX_TRACKSELECT                    As Long = &H8&
Private Const LVS_EX_HEADERDRAGDROP                 As Long = &H10&
Private Const LVS_EX_FULLROWSELECT                  As Long = &H20&
Private Const LVS_EX_ONECLICKACTIVATE               As Long = &H40&
Private Const LVS_EX_TWOCLICKACTIVATE               As Long = &H80&
Private Const LVS_EX_FLATSB                         As Long = &H100&
Private Const LVS_EX_REGIONAL                       As Long = &H200&
Private Const LVS_EX_INFOTIP                        As Long = &H400&
Private Const LVS_EX_UNDERLINEHOT                   As Long = &H800&
Private Const LVS_EX_UNDERLINECOLD                  As Long = &H1000&
Private Const LVS_EX_MULTIWORKAREAS                 As Long = &H2000&
Private Const LVS_EX_LABELTIP                       As Long = &H4000&
Private Const LVS_EX_BORDERSELECT                   As Long = &H8000&
Private Const LVS_EX_DOUBLEBUFFER                   As Long = &H10000
Private Const LVS_EX_HIDELABELS                     As Long = &H20000
Private Const LVS_EX_SINGLEROW                      As Long = &H40000
Private Const LVS_EX_SNAPTOGRID                     As Long = &H80000
Private Const LVS_EX_SIMPLESELECT                   As Long = &H100000

Private Const LVS_ALIGNTOP                          As Long = &H0
Private Const LVS_TYPEMASK                          As Long = &H3
Private Const LVS_SINGLESEL                         As Long = &H4
Private Const LVS_SHOWSELALWAYS                     As Long = &H8
Private Const LVS_SORTASCENDING                     As Long = &H10
Private Const LVS_SORTDESCENDING                    As Long = &H20
Private Const LVS_SHAREIMAGELISTS                   As Long = &H40
Private Const LVS_NOLABELWRAP                       As Long = &H80
Private Const LVS_AUTOARRANGE                       As Long = &H100
Private Const LVS_EDITLABELS                        As Long = &H200
Private Const LVS_ALIGNLEFT                         As Long = &H800
Private Const LVS_ALIGNMASK                         As Long = &HC00
Private Const LVS_OWNERDATA                         As Long = &H1000
Private Const LVS_NOSCROLL                          As Long = &H2000
Private Const LVS_TYPESTYLEMASK                     As Long = &HFC00
Private Const LVS_OWNERDRAWFIXED                    As Long = &H400
Private Const LVS_NOCOLUMNHEADER                    As Long = &H4000
Private Const LVS_NOSORTHEADER                      As Long = &H8000
Private Const LVSCW_AUTOSIZE                        As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER              As Long = -2

Private Const LVSIL_NORMAL                          As Long = 0
Private Const LVSIL_SMALL                           As Long = 1
Private Const LVSIL_STATE                           As Long = 2

Private Const SW_SHOW                               As Long = 5
Private Const GW_CHILD                              As Long = 5
Private Const SWP_NOMOVE                            As Long = &H2
Private Const SWP_NOSIZE                            As Long = &H1
Private Const SWP_NOOWNERZORDER                     As Long = &H200
Private Const SWP_NOZORDER                          As Long = &H4
Private Const SWP_FRAMECHANGED                      As Long = &H20

Private Const WC_LISTVIEW                           As String = "SysListView32"

Private Const WM_SETFONT                            As Long = &H30
Private Const WM_SETFOCUS                           As Long = &H7
Private Const WM_NOTIFY                             As Long = &H4E
Private Const WM_KEYDOWN                            As Long = &H100
Private Const WM_KEYUP                              As Long = &H101
Private Const WM_CHAR                               As Long = &H102
Private Const WM_MOUSEMOVE                          As Long = &H200
Private Const WM_LBUTTONUP                          As Long = &H202
Private Const WM_LBUTTONDOWN                        As Long = &H201
Private Const WM_RBUTTONDOWN                        As Long = &H204
Private Const WM_RBUTTONUP                          As Long = &H205
Private Const WM_MBUTTONDOWN                        As Long = &H207
Private Const WM_MBUTTONUP                          As Long = &H208
Private Const WM_TIMER                              As Long = &H113&

Private Const WS_EX_TOPMOST                         As Long = &H8&
Private Const WS_EX_WINDOWEDGE                      As Long = &H100&
Private Const WS_EX_CLIENTEDGE                      As Long = &H200&
Private Const WS_EX_STATICEDGE                      As Long = &H20000
Private Const WS_TABSTOP                            As Long = &H10000
Private Const WS_THICKFRAME                         As Long = &H40000
Private Const WS_BORDER                             As Long = &H800000
Private Const WS_DISABLED                           As Long = &H8000000
Private Const WS_VISIBLE                            As Long = &H10000000
Private Const WS_CHILD                              As Long = &H40000000

Public Enum eStyle
    [eReport] = LVS_REPORT
    [eIcon] = LVS_ICON
    [eSmallIcon] = LVS_SMALLICON
    [eList] = LVS_LIST
End Enum

Public Enum eBorderStyle
    [bLine] = 0
    [bThin] = 1
    [bThick] = 2
End Enum

Public Enum eColumnAutosize
    [cItem] = LVSCW_AUTOSIZE
    [cHeader] = LVSCW_AUTOSIZE_USEHEADER
End Enum

Public Enum eColumnAlign
    [cLeft] = HDF_LEFT
    [cRight] = HDF_RIGHT
    [ccEnter] = HDF_CENTER
End Enum

Public Enum LVM_SETITEMCOUNT_lParam
    LVSICF_NOINVALIDATEALL = &H1
    LVSICF_NOSCROLL = &H2
End Enum

Public Enum TT_Notifications
    TTN_FIRST = -520&
    TTN_LAST = -549&
    TTN_GETDISPINFO = (TTN_FIRST - 0)
End Enum

Public Enum LISTVIEW_MESSAGES
    'LVM_FIRST = &H1000
    LVM_SETITEMCOUNT = (LVM_FIRST + 47)
    LVM_GETITEMRECT = (LVM_FIRST + 14)
    LVM_SCROLL = (LVM_FIRST + 20)
    LVM_GETTOPINDEX = (LVM_FIRST + 39)
    LVM_HITTEST = (LVM_FIRST + 18)
    LVM_DELETEALLITEMS = (LVM_FIRST + 9)
    LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
    LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
End Enum

Public Enum LV_ITEM_mask
    LVIF_TEXT = &H1
    LVIF_IMAGE = &H2
    LVIF_PARAM = &H4
    LVIF_STATE = &H8
    LVIF_INDENT = &H10
    LVIF_NORECOMPUTE = &H800
    LVIF_DI_SETITEM = &H1000
End Enum

Public Enum LV_ITEM_state
    LVIS_FOCUSED = &H1
    LVIS_SELECTED = &H2
    LVIS_CUT = &H4
    LVIS_DROPHILITED = &H8
    LVIS_OVERLAYMASK = &HF00
    LVIS_STATEIMAGEMASK = &HF000
    LVIS_ALL = LVIS_FOCUSED Or LVIS_SELECTED Or LVIS_CUT Or LVIS_DROPHILITED Or LVIS_OVERLAYMASK Or LVIS_STATEIMAGEMASK
End Enum

Public Enum LVNotifications
    LVN_FIRST = -100&
    LVN_LAST = -199&
    LVN_ITEMCHANGING = (LVN_FIRST - 0)
    LVN_ITEMCHANGED = (LVN_FIRST - 1)
    LVN_INSERTITEM = (LVN_FIRST - 2)
    LVN_DELETEITEM = (LVN_FIRST - 3)
    LVN_DELETEALLITEMS = (LVN_FIRST - 4)
    LVN_COLUMNCLICK = (LVN_FIRST - 8)
    LVN_BEGINDRAG = (LVN_FIRST - 9)
    LVN_BEGINRDRAG = (LVN_FIRST - 11)
    LVN_ODCACHEHINT = (LVN_FIRST - 13)
    LVN_ITEMACTIVATE = (LVN_FIRST - 14)
    LVN_ODSTATECHANGED = (LVN_FIRST - 15)
    LVN_BEGINLABELEDIT = (LVN_FIRST - 5)
    LVN_ENDLABELEDIT = (LVN_FIRST - 6)
    LVN_GETDISPINFO = (LVN_FIRST - 50)
    LVN_SETDISPINFO = (LVN_FIRST - 51)
    LVN_ODFINDITEM = (LVN_FIRST - 52)
    LVN_KEYDOWN = (LVN_FIRST - 55)
    LVN_MARQUEEBEGIN = (LVN_FIRST - 56)
End Enum

Private Type RECT
    left                                            As Long
    top                                             As Long
    right                                           As Long
    bottom                                          As Long
End Type

Private Type POINTAPI
    x                                               As Long
    y                                               As Long
End Type

Private Type LVITEM
    mask                                            As Long
    iItem                                           As Long
    iSubItem                                        As Long
    State                                           As Long
    stateMask                                       As Long
    pszText                                         As String
    cchTextMax                                      As Long
    iImage                                          As Long
    lParam                                          As Long
    iIndent                                         As Long
End Type

Private Type LVHITTESTINFO
    pt                                              As POINTAPI
    flags                                           As Long
    iItem                                           As Long
    iSubItem                                        As Long
End Type

Private Type LVCOLUMN
    mask                                            As Long
    fmt                                             As Long
    cx                                              As Long
    pszText                                         As String
    cchTextMax                                      As Long
    iSubItem                                        As Long
    iImage                                          As Long
    iOrder                                          As Long
End Type


Private Type LVCOLUMNLP
    mask                                            As Long
    fmt                                             As Long
    cx                                              As Long
    pszText                                         As Long
    cchTextMax                                      As Long
    iSubItem                                        As Long
    iImage                                          As Long
    iOrder                                          As Long
End Type

Private Type HDITEM
    mask                                            As Long
    cxy                                             As Long
    pszText                                         As String
    hbm                                             As Long
    cchTextMax                                      As Long
    fmt                                             As Long
    lParam                                          As Long
    iImage                                          As Long
    iOrder                                          As Long
End Type

Private Type NMHDR
    hwndFrom                                        As Long
    idfrom                                          As Long
    code                                            As Long
End Type

Private Type NMCUSTOMDRAWINFO
    hdr                                             As NMHDR
    dwDrawStage                                     As Long
    hdc                                             As Long
    rc                                              As RECT
    dwItemSpec                                      As Long
    iItemState                                      As Long
    lItemLParam                                     As Long
End Type

Private Type NMLVCUSTOMDRAW
    nmcmd                                           As NMCUSTOMDRAWINFO
    clrText                                         As Long
    clrTextBk                                       As Long
    'iSubItem As Integer
End Type

Private Type LV_ITEM
    mask                                            As LV_ITEM_mask
    iItem                                           As Long
    iSubItem                                        As Long
    State                                           As LV_ITEM_state
    stateMask                                       As Long
    pszText                                         As Long
    cchTextMax                                      As Long
    iImage                                          As Long
    lParam                                          As Long
    iIndent                                         As Long
End Type

Private Type LV_DISPINFO
    hdr                                             As NMHDR
    Item                                            As LV_ITEM
End Type

Private Type NM_LISTVIEW
    hdr                                             As NMHDR
    iItem                                           As Long
    iSubItem                                        As Long
    uNewState                                       As Long
    uOldState                                       As Long
    uChanged                                        As Long
    ptAction                                        As POINTAPI
    lParam                                          As Long
End Type

Private Type LOGFONT
    lfHeight                                        As Long
    lfWidth                                         As Long
    lfEscapement                                    As Long
    lfOrientation                                   As Long
    lfWeight                                        As Long
    lfItalic                                        As Byte
    lfUnderline                                     As Byte
    lfStrikeOut                                     As Byte
    lfCharSet                                       As Byte
    lfOutPrecision                                  As Byte
    lfClipPrecision                                 As Byte
    lfQuality                                       As Byte
    lfPitchAndFamily                                As Byte
    lfFaceName(32)                                  As Byte
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal x As Long, _
                                                                              ByVal y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hWndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function lstrcpyFromPointer Lib "kernel32" Alias "lstrcpyA" (ByVal lpDest As String, _
                                                                             ByVal lpSource As Long) As Long

Private Declare Function lstrcpyToPointer Lib "kernel32" Alias "lstrcpyA" (ByVal lpDest As Long, _
                                                                           ByVal lpSource As String) As Long

Private Declare Function lstrlenFromPointer Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function ImageList_Create Lib "Comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Add Lib "Comctl32" (ByVal hImageList As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function ImageList_SetBkColor Lib "Comctl32" (ByVal hImageList As Long, _
                                                              ByVal clrBk As Long) As Long

Private Declare Function ImageList_AddMasked Lib "Comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_AddIcon Lib "Comctl32" (ByVal hImageList As Long, _
                                                           ByVal hIcon As Long) As Long

Private Declare Function ImageList_Destroy Lib "Comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_Draw Lib "Comctl32" (ByVal hImageList As Long, _
                                                        ByVal lIndex As Long, _
                                                        ByVal hdc As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long, _
                                                        ByVal fStyle As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, _
                                                           ByVal HPALETTE As Long, _
                                                           pccolorref As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Long, _
                                                      ByVal bErase As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                  pSrc As Any, _
                                                                  ByVal lByteLen As Long)

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Public Event eHItemClick(lItem As Long)
Public Event eHItemCheck(lItem As Long)
Public Event eHColumnClick(Column As Long)
'Public Event eHMouseMove(Button As Long, Shift As Long, x As Single, y As Single)
Public Event eHIndirect(ByVal iItem As Long, ByVal iSubItem As Long, ByVal fMask As Long, sText As String, hImage As Long)
Public Event eHErrCond(sRtn As String, lErr As Long)

'/* note: organize data types in order, and test operand alignment
Private m_bCheckBoxes                               As Boolean
Private m_bSortDirection                            As Boolean
Private m_bFirstItem                                As Boolean
Private m_bSubImages                                As Boolean
Private m_bFullRowSelect                            As Boolean
Private m_bGridLines                                As Boolean
Private m_bDragDrop                                 As Boolean
Private m_bHeaderFixed                              As Boolean
Private m_bHeaderFlat                               As Boolean
Private m_bHeaderHide                               As Boolean
Private m_bHideSelection                            As Boolean
Private m_bLabelTips                                As Boolean
Private m_bLabelEdit                                As Boolean
Private m_bMultiSelect                              As Boolean
Private m_bOneClick                                 As Boolean
Private m_bScrollFlat                               As Boolean
Private m_bUnderlineHot                             As Boolean
Private m_bCheckInit                                As Boolean
Private m_bIndirectMode                             As Boolean
Private m_bCustomHeader                             As Boolean
Private m_bSorted                                   As Boolean
Private m_bUseSorted                                As Boolean
Private m_bSubItemImage                             As Boolean
Private m_lParentHwnd                               As Long
Private m_lLVHwnd                                   As Long
Private m_lHdrHwnd                                  As Long
Private m_lImlHdHndl                                As Long
Private m_lImlSmallHndl                             As Long
Private m_lImlLargeHndl                             As Long
Private m_lItemsCnt                                 As Long
Private m_lFont                                     As Long
Private m_lPtrIdx                                   As Long
Private m_lCheckState()                             As Long
Private m_lStrctPtr                                 As Long
Private m_lPtr()                                    As Long
Private c_PtrMem                                    As Collection
Private tHLIStc()                                   As HLIStc
Private m_oHdrBkClr                                 As OLE_COLOR
Private m_oHdrForeClr                               As OLE_COLOR
Private WithEvents m_oFont                          As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_ViewMode                                  As eStyle
Private m_BorderStyle                               As eBorderStyle
Private m_GSubclass                                 As MGSubclass


Private Sub Class_Initialize()

    Set m_GSubclass = New MGSubclass
    ReDim tHLIStc(0)
    m_oHdrBkClr = GetSysColor(m_oHdrBkClr And &H1F&)
    m_oHdrForeClr = GetSysColor(vbWindowText And &H1F&)

End Sub

'* Name           : [get] p_StructPntr
'* Purpose        : retrieve pointer to the data structure
'* Inputs         : none
'* Outputs        : long
'*********************************************
Public Property Get p_StructPtr() As Long
    p_StructPtr = m_lStrctPtr
End Property

'* Name           : [let] p_StructPntr
'* Purpose        : add pointer to the data structure
'* Inputs         : long
'* Outputs        : none
'*********************************************
Public Property Let p_StructPtr(ByVal PropVal As Long)
    m_lStrctPtr = PropVal
End Property

'* Name           : LoadArray
'* Purpose        : load data structure
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Function LoadArray() As Boolean

Dim lSpr    As Long

On Error GoTo Handler

    Set c_PtrMem = New Collection
    '/* initialize local struct
    ReDim tHLIStc(0)
    '/* copy the structure from the pointer
    CopyMemBr ByVal GetAVPtr(tHLIStc), m_lStrctPtr, &H4
    c_PtrMem.Add m_lStrctPtr, "tHLIStc"
    LoadArray = True
    
Handler:
    On Error GoTo 0
    
End Function

'* Name           : CreateList
'* Purpose        : initialize the listview
'* Inputs         : style, width, height, parent hwnd, app instance
'* Outputs        : boolean
'*********************************************
Public Function CreateList(Style As eStyle, _
                           ByVal lWidth As Long, _
                           ByVal lHeight As Long, _
                           ByVal lCntHnd As Long, _
                           ByVal lAppInst As Long) As Boolean


Dim lLVStyle    As Long
Dim lExStyle    As Long

On Error GoTo Handler
    
    '/* destroy existing
    DestroyList
    '/* initial style flags including LVS_OWNERDATA
    '/* this tells the list that all data will be
    '/* managed externally
    lLVStyle = WS_CHILD Or WS_BORDER Or WS_VISIBLE Or Style Or LVS_SORTASCENDING Or LVS_OWNERDATA Or _
        LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_SINGLESEL Or WS_TABSTOP Or Style
    '/* default container style
    lExStyle = GetWindowLong(lCntHnd, GWL_EXSTYLE) And Not WS_EX_CLIENTEDGE
    '/* create the list
    m_lLVHwnd = CreateWindowEx(lExStyle, WC_LISTVIEW, vbNullString, lLVStyle, _
    0, 0, lWidth, lHeight, lCntHnd, 0, lAppInst, ByVal 0&)

    '/* subclass the list and parent WM_NOTIFY messages
    '/* control callback data is reflected from parent control
    If Not m_lLVHwnd = 0 Then
        m_lParentHwnd = lCntHnd
        List_Attatch m_lLVHwnd
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("CreateList", Err.Number)

End Function

'> Columns/Headers
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : ColumnAdd
'* Purpose        : create column headers
'* Inputs         : column index, text, width, text alignment, icon index
'* Outputs        : boolean
'*********************************************
Public Function ColumnAdd(ByVal lIndex As Long, _
                          ByVal sText As String, _
                          ByVal lWidth As Long, _
                          Optional ByVal eAlign As eColumnAlign = [cLeft], _
                          Optional ByVal lIcon As Long = -1) As Boolean


Dim bFirst  As Boolean
Dim uLVC    As LVCOLUMN
Dim uHDI    As HDITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    bFirst = (Me.p_ColumnCount = 0)
    ColumnAdd = (SendMessage(m_lLVHwnd, LVM_INSERTCOLUMN, lIndex, uLVC) > -1)
    If ColumnAdd Then
        If bFirst Then
            m_lHdrHwnd = HeaderHwnd()
            SendMessageLong m_lHdrHwnd, HDM_SETIMAGELIST, 0, m_lImlHdHndl
        End If
        With uHDI
            .pszText = sText + vbNullChar
            .cchTextMax = Len(sText) + 1
            .cxy = lWidth
            .iImage = lIcon
            .fmt = HDF_STRING Or eAlign * -(lIndex <> 0) Or HDF_IMAGE * -(lIcon > -1) Or HDF_BITMAP_ON_RIGHT
            .mask = HDI_TEXT Or HDI_WIDTH Or HDI_IMAGE Or HDI_FORMAT
        End With
        SendMessage m_lHdrHwnd, HDM_SETITEM, lIndex, uHDI
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnAdd", Err.Number)

End Function

'* Name           : [get] p_ColumnText
'* Purpose        : get a columns heading
'* Inputs         : column index
'* Outputs        : string
'*********************************************
Public Property Get p_ColumnText(ByVal lColumn As Long) As String

Dim uLVC            As LVCOLUMNLP
Dim aText(261)      As Byte

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
        .pszText = VarPtr(aText(0))
        .cchTextMax = UBound(aText)
        .mask = LVCF_TEXT
    End With
    SendMessage m_lLVHwnd, LVM_GETCOLUMN, lColumn, uLVC
    p_ColumnText = left$(StrConv(aText(), vbUnicode), uLVC.cchTextMax)

End Property

'* Name           : [let] p_ColumnText
'* Purpose        : change a columns heading
'* Inputs         : column index, text
'* Outputs        : none
'*********************************************
Public Property Let p_ColumnText(ByVal lColumn As Long, _
                                 ByVal sText As String)

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
        .pszText = sText & vbNullChar
        .cchTextMax = Len(sText) + 1
        .mask = LVCF_TEXT
    End With
    SendMessage m_lLVHwnd, LVM_SETCOLUMN, lColumn, uLVC

End Property

'* Name           : [get] p_ColumnWidth
'* Purpose        : retrieve a columns length
'* Inputs         : column index
'* Outputs        : long
'*********************************************
Public Property Get p_ColumnWidth(ByVal lColumn As Long) As Long

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    p_ColumnWidth = SendMessageLong(m_lLVHwnd, LVM_GETCOLUMNWIDTH, lColumn, 0)

End Property

'* Name           : [let] p_ColumnWidth
'* Purpose        : change a columns length
'* Inputs         : column index, width
'* Outputs        : long
'*********************************************
Public Property Let p_ColumnWidth(ByVal lColumn As Long, _
                                  ByVal lWidth As Long)

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    SendMessageLong m_lLVHwnd, LVM_SETCOLUMNWIDTH, lColumn, lWidth

End Property

'* Name           : [get] p_ColumnAlign
'* Purpose        : retieve a columns text alignment
'* Inputs         : column index
'* Outputs        : eColumnAlign
'*********************************************
Public Property Get p_ColumnAlign(ByVal lColumn As Long) As eColumnAlign

Const lMask     As Long = &H3
Dim uLVC        As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    uLVC.mask = LVCF_FMT
    SendMessage m_lLVHwnd, LVM_GETCOLUMN, lColumn, uLVC
    p_ColumnAlign = (lMask And uLVC.fmt)

End Property

'* Name           : [let] p_ColumnAlign
'* Purpose        : change a columns text alignment
'* Inputs         : column index
'* Outputs        : column index, eColumnAlign
'*********************************************
Public Property Let p_ColumnAlign(ByVal lColumn As Long, _
                                  ByVal eAlign As eColumnAlign)

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uLVC
        .fmt = eAlign * -(Not lColumn = 0)
        .mask = LVCF_FMT
    End With
    SendMessage m_lLVHwnd, LVM_SETCOLUMN, lColumn, uLVC

End Property

'* Name           : [get] p_ColumnIcon
'* Purpose        : retieve header icon index
'* Inputs         : column index
'* Outputs        : image index
'*********************************************
Public Property Get p_ColumnIcon(ByVal lColumn As Long) As Long

Dim uLVC    As LVCOLUMN

    If m_lLVHwnd = 0 Then Exit Property
    uLVC.mask = LVCF_IMAGE
    SendMessage m_lLVHwnd, LVM_GETCOLUMN, lColumn, uLVC
    p_ColumnIcon = uLVC.iImage

End Property

'* Name           : [let] p_ColumnIcon
'* Purpose        : change header icon
'* Inputs         : column index, icon index
'* Outputs        : none
'*********************************************
Public Property Let p_ColumnIcon(ByVal lColumn As Long, _
                                 ByVal lIcon As Long)


Const lMask     As Long = &H3
Dim lAlign      As Long
Dim uHDI        As HDITEM

    If m_lLVHwnd = 0 Or m_lHdrHwnd = 0 Then Exit Property
    With uHDI
        .mask = HDI_FORMAT
        SendMessage m_lHdrHwnd, HDM_GETITEM, lColumn, uHDI
        lAlign = lMask And .fmt
        .iImage = lIcon
        .fmt = HDF_STRING Or lAlign Or HDF_IMAGE * -(lIcon > -1 And m_lImlHdHndl <> 0) Or HDF_BITMAP_ON_RIGHT
        .mask = HDI_IMAGE * -(lIcon > -1) Or HDI_FORMAT
    End With
    SendMessage m_lHdrHwnd, HDM_SETITEM, lColumn, uHDI

End Property

'* Name           : [get] p_ColumnCount
'* Purpose        : retieve column count
'* Inputs         : none
'* Outputs        : column count
'*********************************************
Public Property Get p_ColumnCount() As Long

    If m_lLVHwnd = 0 Then Exit Property
    p_ColumnCount = SendMessageLong(HeaderHwnd(), HDM_GETITEMCOUNT, 0, 0)

End Property

'* Name           : HeaderHwnd
'* Purpose        : return the column header handle
'* Inputs         : none
'* Outputs        : none
'*********************************************
Private Function HeaderHwnd() As Long

    If m_lLVHwnd = 0 Then Exit Function
    HeaderHwnd = SendMessageLong(m_lLVHwnd, LVM_GETHEADER, 0, 0)

End Function

'* Name           : [get] p_HeaderColor
'* Purpose        : return the header color
'* Inputs         : none
'* Outputs        : ole color
'*********************************************
Public Property Get p_HeaderColor() As OLE_COLOR
    p_HeaderColor = m_oHdrBkClr
End Property

'* Name           : [let] p_HeaderColor
'* Purpose        : change the header color
'* Inputs         : ole color
'* Outputs        : none
'*********************************************
Public Property Let p_HeaderColor(PropVal As OLE_COLOR)
    m_oHdrBkClr = PropVal
End Property

'* Name           : [get] p_HeaderForeColor
'* Purpose        : return the header forecolor
'* Inputs         : none
'* Outputs        : ole color
'*********************************************
Public Property Get p_HeaderForeColor() As OLE_COLOR
    p_HeaderForeColor = m_oHdrForeClr
End Property

'* Name           : [let] p_HeaderForeColor
'* Purpose        : change the header forecolor
'* Inputs         : ole color
'* Outputs        : none
'*********************************************
Public Property Let p_HeaderForeColor(PropVal As OLE_COLOR)
    m_oHdrForeClr = PropVal
End Property

'* Name           : [get] p_HeaderCustom
'* Purpose        : return the custom header status
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_HeaderCustom() As Boolean
    p_HeaderCustom = m_bCustomHeader
End Property

'* Name           : [let] p_HeaderCustom
'* Purpose        : change the custom header status
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_HeaderCustom(PropVal As Boolean)
    m_bCustomHeader = PropVal
End Property

'* Name           : ColumnClear
'* Purpose        : remove all columns
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Function ColumnClear() As Boolean

Dim lCt As Long

    For lCt = p_ColumnCount To 0 Step -1
        ColumnRemove lCt
    Next lCt
    
End Function

'* Name           : ColumnRemove
'* Purpose        : remove a column
'* Inputs         : column index
'* Outputs        : boolean
'*********************************************
Public Function ColumnRemove(ByVal lColumn As Long) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ColumnRemove = CBool(SendMessageLong(m_lLVHwnd, LVM_DELETECOLUMN, lColumn, 0))
    If Me.p_ColumnCount = 0 Then m_lHdrHwnd = 0

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnRemove", Err.Number)

End Function

'* Name           : ColumnAutosize
'* Purpose        : autosize columns
'* Inputs         : column index, size constant
'* Outputs        : boolean
'*********************************************
Public Function ColumnAutosize(ByVal lColumn As Long, _
                               Optional ByVal AutosizeType As eColumnAutosize = [cItem]) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ColumnAutosize = CBool(SendMessageLong(m_lLVHwnd, LVM_SETCOLUMNWIDTH, lColumn, AutosizeType))

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ColumnAutosize", Err.Number)

End Function

'* Name           : p_Count
'* Purpose        : [get] item count
'* Inputs         : none
'* Outputs        : long
'*********************************************
Public Property Get p_Count() As Long

    If m_lLVHwnd Then
        p_Count = SendMessageLong(m_lLVHwnd, LVM_GETITEMCOUNT, 0, 0)
    End If

End Property

'>  ImageLists
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : InitImlHeader
'* Purpose        : initialize header imagelist
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Function InitImlHeader() As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlHeader
    m_lImlHdHndl = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, 0, 0)
    InitImlHeader = (Not m_lImlHdHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("InitImlHeader", Err.Number)
    
End Function

'* Name           : ImlHeaderAddBmp
'* Purpose        : add a bitmap to header iml
'* Inputs         : bmp hndl, mask color
'* Outputs        : Long
'*********************************************
Public Function ImlHeaderAddBmp(ByVal lBitmap As Long, _
                                Optional ByVal lMaskColor As Long = CLR_NONE) As Long

On Error GoTo Handler

    If m_lImlHdHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlHeaderAddBmp = ImageList_AddMasked(m_lImlHdHndl, lBitmap, lMaskColor)
    Else
        ImlHeaderAddBmp = ImageList_Add(m_lImlHdHndl, lBitmap, 0)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlHeaderAddBmp", Err.Number)

End Function

'* Name           : ImlHeaderAddIcon
'* Purpose        : add an icon to header iml
'* Inputs         : icon handle
'* Outputs        : Long
'*********************************************
Public Function ImlHeaderAddIcon(ByVal lIcon As Long) As Long

On Error GoTo Handler

    If m_lImlHdHndl = 0 Then Exit Function
    ImlHeaderAddIcon = ImageList_AddIcon(m_lImlHdHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlHeaderAddIcon", Err.Number)

End Function

'* Name           : DestroyImlHeader
'* Purpose        : destroy header image list
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Private Function DestroyImlHeader() As Boolean

On Error GoTo Handler

    If m_lImlHdHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlHdHndl) Then
        DestroyImlHeader = True
        m_lImlHdHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlHeader", Err.Number)

End Function

'* Name           : InitImlSmall
'* Purpose        : initialize smallicons image list
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Function InitImlSmall() As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlSmall
    m_lImlSmallHndl = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, 0, 0)
    SendMessageLong m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, m_lImlSmallHndl
    InitImlSmall = (Not m_lImlSmallHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("InitImlSmall", Err.Number)

End Function

'* Name           : ImlSmallAddBmp
'* Purpose        : add bmp to small image iml
'* Inputs         : bitmap handle, mask color
'* Outputs        : Long
'*********************************************
Public Function ImlSmallAddBmp(ByVal lBitmap As Long, _
                               Optional ByVal lMaskColor As Long = CLR_NONE) As Long

On Error GoTo Handler

    If m_lImlSmallHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlSmallAddBmp = ImageList_AddMasked(m_lImlSmallHndl, lBitmap, lMaskColor)
    Else
        ImlSmallAddBmp = ImageList_Add(m_lImlSmallHndl, lBitmap, 0)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlSmallAddBmp", Err.Number)

End Function

'* Name           : ImlSmallAddIcon
'* Purpose        : add icon to small image iml
'* Inputs         : icon handle
'* Outputs        : Long
'*********************************************
Public Function ImlSmallAddIcon(ByVal lIcon As Long) As Long

On Error GoTo Handler

    If m_lImlSmallHndl = 0 Then Exit Function
    ImlSmallAddIcon = ImageList_AddIcon(m_lImlSmallHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlSmallAddIcon", Err.Number)

End Function

'* Name           : DestroyImlSmall
'* Purpose        : destroy small icons image list
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Private Function DestroyImlSmall() As Boolean

On Error GoTo Handler

    If m_lImlSmallHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlSmallHndl) Then
        DestroyImlSmall = True
        m_lImlSmallHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlSmall", Err.Number)

End Function

'* Name           : InitImlLarge
'* Purpose        : initialize large icons image list
'* Inputs         : image height/width
'* Outputs        : boolean
'*********************************************
Public Function InitImlLarge(Optional ByVal lImgWidth As Long = 32, _
                             Optional ByVal lImgHeight As Long = 32) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    DestroyImlLarge
    m_lImlLargeHndl = ImageList_Create(lImgWidth, lImgHeight, ILC_COLOR32 Or ILC_MASK, 0, 0)
    SendMessageLong m_lLVHwnd, LVM_SETIMAGELIST, LVSIL_NORMAL, m_lImlLargeHndl
    InitImlLarge = (Not m_lImlLargeHndl = 0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlLargeAddIcon", Err.Number)

End Function

'* Name           : ImlLargeAddBmp
'* Purpose        : add bmp to large image iml
'* Inputs         : bitmap handle, mask color
'* Outputs        : Long
'*********************************************
Public Function ImlLargeAddBmp(ByVal lBitmap As Long, _
                               Optional ByVal lMaskColor As Long = CLR_NONE) As Long

On Error GoTo Handler

    If m_lImlLargeHndl = 0 Then Exit Function
    If Not lMaskColor = CLR_NONE Then
        ImlLargeAddBmp = ImageList_AddMasked(m_lImlLargeHndl, lBitmap, lMaskColor)
    Else
        ImlLargeAddBmp = ImageList_Add(m_lImlLargeHndl, lBitmap, 0)
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlLargeAddBmp", Err.Number)

End Function

'* Name           : ImlLargeAddIcon
'* Purpose        : add icon to large image iml
'* Inputs         : icon handle
'* Outputs        : Long
'*********************************************
Public Function ImlLargeAddIcon(ByVal lIcon As Long) As Long

On Error GoTo Handler

    If m_lImlLargeHndl = 0 Then Exit Function
    ImlLargeAddIcon = ImageList_AddIcon(m_lImlLargeHndl, lIcon)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ImlLargeAddIcon", Err.Number)

End Function

'* Name           : DestroyImlLarge
'* Purpose        : destroy large icons image list
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Private Function DestroyImlLarge() As Boolean

On Error GoTo Handler

    If m_lImlLargeHndl = 0 Then Exit Function
    If ImageList_Destroy(m_lImlLargeHndl) Then
        DestroyImlLarge = True
        m_lImlLargeHndl = 0
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyImlLarge", Err.Number)

End Function

'>  Items
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : ItemAdd
'* Purpose        : add a single item to the list
'* Inputs         : item index, text, indent, icon
'* Outputs        : boolean
'*********************************************
Public Function ItemAdd(ByVal lItem As Long, _
                        ByVal sText As String, _
                        ByVal lIcon As Long, _
                        ByRef sSubItem() As String) As Boolean

Dim lC  As Long
Dim lLb As Long
Dim lUb As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    '/* redimn arrays and add items
    With tHLIStc(0)
        If ArrayExists(.Item) Then
            lLb = LBound(.Item)
            lUb = UBound(.Item) + 1
        Else
            lLb = 0
            lUb = 0
        End If
        On Error Resume Next
        ReDim Preserve .Item(lLb To lUb)
        .Item(lUb) = sText
        ReDim Preserve .lIcon(lLb To lUb)
        .lIcon(lUb) = lIcon
        ReDim Preserve .SubItem1(lLb To lUb)
        .SubItem1(lUb) = sSubItem(0)
        ReDim Preserve .SubItem2(lLb To lUb)
        .SubItem2(lUb) = sSubItem(1)
    End With
    
    '/* redimension and trigger list
    SetItemCount lUb + 1

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemAdd", Err.Number)

End Function

'* Name           : ItemRemove
'* Purpose        : remove an item from the list
'* Inputs         : item index
'* Outputs        : boolean
'*********************************************
Public Function ItemRemove(ByVal lItem As Long) As Boolean

Dim lLb         As Long
Dim lUb         As Long
Dim lC          As Long
Dim lP          As Long
Dim lS          As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    With tHLIStc(0)
        ArrayRemString .Item, lItem
        ArrayRemString .SubItem1, lItem
        ArrayRemString .SubItem2, lItem
        ArrayRemLong .lIcon, lItem
    End With
    
    SetItemCount UBound(tHLIStc(0).Item) + 1

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemRemove", Err.Number)

End Function

'* Name           : ItemEnsureVisible
'* Purpose        : move to item index
'* Inputs         : item index
'* Outputs        : boolean
'*********************************************
Public Function ItemEnsureVisible(ByVal lItem As Long) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function
    ItemEnsureVisible = CBool(SendMessageLong(m_lLVHwnd, LVM_ENSUREVISIBLE, lItem, 0))

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("ItemEnsureVisible", Err.Number)

End Function

'* Name           : [get] p_ItemsSorted
'* Purpose        : return sorted mode status
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_ItemsSorted() As Boolean

    p_ItemsSorted = m_bUseSorted

End Property

'* Name           : [let] p_ItemsSorted
'* Purpose        : change sorted mode status
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_ItemsSorted(PropVal As Boolean)

    m_bUseSorted = PropVal

End Property

'* Name           : [get] p_IndirectMode
'* Purpose        : return indirect mode status
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_IndirectMode() As Boolean

    p_IndirectMode = m_bIndirectMode

End Property

'* Name           : [let] p_IndirectMode
'* Purpose        : change indirect mode status
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_IndirectMode(PropVal As Boolean)

    m_bIndirectMode = PropVal

End Property

'* Name           : [get] p_ItemText
'* Purpose        : return item text
'* Inputs         : item index
'* Outputs        : string
'*********************************************
Public Property Get p_ItemText(ByVal lItem As Long) As String

Dim uLVI   As LVITEM
Dim a(261) As Byte
Dim lLen   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .pszText = VarPtr(a(0))
        .cchTextMax = UBound(a)
    End With
    
    lLen = SendMessage(m_lLVHwnd, LVM_GETITEMTEXT, lItem, uLVI)
    p_ItemText = left$(StrConv(a(), vbUnicode), lLen)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemText", Err.Number)

End Property

'* Name           : [let] p_ItemText
'* Purpose        : change item text
'* Inputs         : item index, string
'* Outputs        : string
'*********************************************
Public Property Let p_ItemText(ByVal lItem As Long, _
                               ByVal sText As String)

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .pszText = sText & vbNullChar
        .cchTextMax = Len(sText) + 1
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMTEXT, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemText", Err.Number)

End Property

'* Name           : [get] p_ItemIcon
'* Purpose        : return icon index
'* Inputs         : item index
'* Outputs        : long
'*********************************************
Public Property Get p_ItemIcon(ByVal lItem As Long) As Long

Dim uLVI As LVITEM

On Error Resume Next

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iItem = lItem
        .mask = LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_GETITEM, 0, uLVI
    p_ItemIcon = uLVI.iImage
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemIcon", Err.Number)

End Property

'* Name           : [let] p_ItemIcon
'* Purpose        : change icon index
'* Inputs         : item index
'* Outputs        : long
'*********************************************
Public Property Let p_ItemIcon(ByVal lItem As Long, _
                               ByVal lIcon As Long)

Dim uLVI As LVITEM

On Error GoTo Handler

'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/listview/macros/listview_setitemstate.asp

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iItem = lItem
        .iImage = lIcon
        .mask = LVIS_OVERLAYMASK
        .State = lItem * 256
       ' LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, 0, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemIcon", Err.Number)

End Property

'* Name           : [get] p_ItemIndent
'* Purpose        : return item indent
'* Inputs         : item index
'* Outputs        : long
'*********************************************
Public Property Get p_ItemIndent(ByVal lItem As Long) As Long

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iItem = lItem
        .mask = LVIF_INDENT
    End With
    SendMessage m_lLVHwnd, LVM_GETITEM, 0, uLVI
    p_ItemIndent = uLVI.iIndent

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemIndent", Err.Number)

End Property

'* Name           : [let] p_ItemIndent
'* Purpose        : change item indent
'* Inputs         : item index, indent
'* Outputs        : none
'*********************************************
Public Property Let p_ItemIndent(ByVal lItem As Long, _
                                 ByVal Indent As Long)

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iItem = lItem
        .iIndent = Indent
        .mask = LVIF_INDENT
    End With
    SendMessage m_lLVHwnd, LVM_SETITEM, 0, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemIndent", Err.Number)

End Property

'* Name           : [get] p_ItemSelected
'* Purpose        : return selected state
'* Inputs         : item index
'* Outputs        : boolean
'*********************************************
Public Property Get p_ItemSelected(ByVal lItem As Long) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    p_ItemSelected = CBool(SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_SELECTED))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemSelected", Err.Number)

End Property

'* Name           : [let] p_ItemSelected
'* Purpose        : select an item
'* Inputs         : item index, indent
'* Outputs        : none
'*********************************************
Public Property Let p_ItemSelected(ByVal lItem As Long, _
                                   ByVal bSelected As Boolean)

Dim uLVI    As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_SELECTED Or -(bSelected And lItem > -1) * LVIS_FOCUSED
        .State = -bSelected * LVIS_SELECTED Or -(lItem > -1) * LVIS_FOCUSED
        .mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemSelected", Err.Number)

End Property

'* Name           : [get] p_ItemFocused
'* Purpose        : return item focused state
'* Inputs         : item index
'* Outputs        : boolean
'*********************************************
Public Property Get p_ItemFocused(ByVal lItem As Long) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    p_ItemFocused = CBool(SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_FOCUSED))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemFocused", Err.Number)

End Property

'* Name           : [let] p_ItemFocused
'* Purpose        : change item focused state
'* Inputs         : item index, focus
'* Outputs        : none
'*********************************************
Public Property Let p_ItemFocused(ByVal lItem As Long, _
                                  ByVal bFocused As Boolean)

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    With uLVI
        .stateMask = LVIS_FOCUSED
        .State = -bFocused * LVIS_FOCUSED
        .mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemFocused", Err.Number)

End Property

'* Name           : [get] p_ItemChecked
'* Purpose        : return item checked state
'* Inputs         : item index
'* Outputs        : boolean
'*********************************************
Public Property Get p_ItemChecked(ByVal lItem As Long) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    p_ItemChecked = ((SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_STATEIMAGEMASK) And &H2000&) = &H2000&)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemChecked", Err.Number)

End Property

'* Name           : [let] p_ItemChecked
'* Purpose        : set item checked state
'* Inputs         : item index, checked
'* Outputs        : none
'*********************************************
Public Property Let p_ItemChecked(ByVal lItem As Long, _
                                  ByVal bChecked As Boolean)

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .stateMask = LVIS_STATEIMAGEMASK
        .State = &H1000& * (1 - bChecked)
        .mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemChecked", Err.Number)

End Property

'* Name           : [get] p_ItemGhosted
'* Purpose        : return item ghosted state
'* Inputs         : item index
'* Outputs        : boolean
'*********************************************
Public Property Get p_ItemGhosted(ByVal lItem As Long) As Boolean

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    p_ItemGhosted = (SendMessageLong(m_lLVHwnd, LVM_GETITEMSTATE, lItem, LVIS_CUT))

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemGhosted", Err.Number)

End Property

'* Name           : [let] p_ItemGhosted
'* Purpose        : change item ghosted state
'* Inputs         : item index
'* Outputs        : none
'*********************************************
Public Property Let p_ItemGhosted(ByVal lItem As Long, _
                                  ByVal bGhosted As Boolean)

Dim uLVI As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .stateMask = LVIS_CUT
        .State = LVIS_CUT * -bGhosted
        .mask = LVIF_STATE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMSTATE, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ItemGhosted", Err.Number)

End Property


'>  SubItems
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : p_SubItemSet
'* Purpose        : change a subitem
'* Inputs         : item index, subitem, text, icon
'* Outputs        : boolean
'*********************************************
Public Function p_SubItemSet(ByVal lItem As Long, _
                             ByVal lSubItem As Long, _
                             ByVal sText As String, _
                             ByVal lIcon As Long) As Boolean

Dim uLV     As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Function 'todo
    With uLV
        .iItem = lItem
        .iSubItem = lSubItem
        .pszText = sText & vbNullChar
        .cchTextMax = Len(sText) + 1
        .iImage = lIcon
        .mask = LVIF_TEXT Or LVIF_IMAGE
    End With
    p_SubItemSet = CBool(SendMessage(m_lLVHwnd, LVM_SETITEM, 0, uLV))

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("p_SubItemSet", Err.Number)

End Function

'* Name           : [get] p_SubItemText
'* Purpose        : retieve subitem text
'* Inputs         : item index, subitem
'* Outputs        : string
'*********************************************
Public Property Get p_SubItemText(ByVal lItem As Long, _
                                  ByVal lSubItem As Long) As String

Dim uLVI        As LVITEM
Dim aText(256)  As Byte
Dim lLen        As Long


On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iSubItem = lSubItem
        .pszText = VarPtr(aText(0))
        .cchTextMax = UBound(aText)
        .mask = LVIF_TEXT
    End With
    lLen = SendMessage(m_lLVHwnd, LVM_GETITEMTEXT, lItem, uLVI)
    p_SubItemText = left$(StrConv(aText(), vbUnicode), lLen)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_SubItemText", Err.Number)

End Property

'* Name           : [let] p_SubItemText
'* Purpose        : change subitem text
'* Inputs         : item index, subitem, text
'* Outputs        : none
'*********************************************
Public Property Let p_SubItemText(ByVal lItem As Long, _
                                  ByVal lSubItem As Long, _
                                  ByVal sText As String)

Dim uLVI    As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iSubItem = lSubItem
        .pszText = sText & vbNullChar
        .cchTextMax = Len(sText) + 1
    End With
    SendMessage m_lLVHwnd, LVM_SETITEMTEXT, lItem, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_SubItemText", Err.Number)

End Property

'* Name           : [get] p_SubItemIcon
'* Purpose        : retrieve subitem icon
'* Inputs         : item index, subitem
'* Outputs        : long
'*********************************************
Public Property Get p_SubItemIcon(ByVal lItem As Long, _
                                  ByVal lSubItem As Long) As Long

Dim uLVI    As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iItem = lItem
        .iSubItem = lSubItem
        .mask = LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_GETITEM, 0, uLVI
    p_SubItemIcon = uLVI.iImage

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_SubItemIcon", Err.Number)

End Property

'* Name           : [let] p_SubItemIcon
'* Purpose        : change subitem icon
'* Inputs         : item index, subitem, icon
'* Outputs        : none
'*********************************************
Public Property Let p_SubItemIcon(ByVal lItem As Long, _
                                  ByVal lSubItem As Long, _
                                  ByVal lIcon As Long)


Dim uLVI    As LVITEM

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property 'todo
    With uLVI
        .iItem = lItem
        .iSubItem = lSubItem
        .iImage = lIcon
        .mask = LVIF_IMAGE
    End With
    SendMessage m_lLVHwnd, LVM_SETITEM, 0, uLVI

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_SubItemIcon", Err.Number)

End Property

'* Name           : [get] p_SubItemImages
'* Purpose        : retrieve subitem icon state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_SubItemImages() As Boolean

    p_SubItemImages = m_bSubItemImage

End Property

'* Name           : [let] p_SubItemImages
'* Purpose        : change subitem icon state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_SubItemImages(ByVal PropVal As Boolean)

    m_bSubItemImage = PropVal

End Property


'>  Listview Properties
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'* Name           : [get] p_BackColor
'* Purpose        : retieve list backcolor
'* Inputs         : none
'* Outputs        : ole color
'*********************************************
Public Property Get p_BackColor() As OLE_COLOR

    If m_lLVHwnd = 0 Then Exit Property
    p_BackColor = SendMessageLong(m_lLVHwnd, LVM_GETBKCOLOR, 0, 0)

End Property

'* Name           : [let] p_BackColor
'* Purpose        : change list backcolor
'* Inputs         : ole color
'* Outputs        : none
'*********************************************
Public Property Let p_BackColor(ByVal PropVal As OLE_COLOR)

Dim lColor As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    OleTranslateColor PropVal, 0, lColor
    SendMessageLong m_lLVHwnd, LVM_SETBKCOLOR, 0, lColor
    SendMessageLong m_lLVHwnd, LVM_SETTEXTBKCOLOR, 0, lColor
    ListRefresh

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_BackColor", Err.Number)

End Property

'* Name           : [get] p_BorderStyle
'* Purpose        : retrive list borderstyle
'* Inputs         : none
'* Outputs        : enum
'*********************************************
Public Property Get p_BorderStyle() As eBorderStyle

    p_BorderStyle = m_BorderStyle

End Property

'* Name           : [let] p_BorderStyle
'* Purpose        : change list borderstyle
'* Inputs         : enum
'* Outputs        : none
'*********************************************
Public Property Let p_BorderStyle(ByVal PropVal As eBorderStyle)

    m_BorderStyle = PropVal
    SetBorderStyle m_lParentHwnd, m_BorderStyle

End Property

'* Name           : [get] p_CheckBoxes
'* Purpose        : retrieve checkbox state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_CheckBoxes() As Boolean

    p_CheckBoxes = m_bCheckBoxes

End Property

'* Name           : [let] p_CheckBoxes
'* Purpose        : change checkbox state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_CheckBoxes(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bCheckBoxes = PropVal
    If m_bCheckBoxes Then
        SetExtendedStyle LVS_EX_CHECKBOXES, 0
    Else
        SetExtendedStyle 0, LVS_EX_CHECKBOXES
    End If
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_CheckBoxes", Err.Number)
    
End Property

'* Name           : [get] p_Font
'* Purpose        : retieve list font
'* Inputs         : none
'* Outputs        : font
'*********************************************
Public Property Get p_Font() As StdFont

On Error GoTo Handler

    Set p_Font = m_oFont

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_Font", Err.Number)

End Property

'* Name           : [let] p_Font
'* Purpose        : change list font
'* Inputs         : font
'* Outputs        : none
'*********************************************
Public Property Set p_Font(ByVal PropVal As StdFont)

Dim uLF     As LOGFONT
Dim lChar   As Long
Dim lHdc    As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    Set m_oFont = PropVal
    With uLF
        For lChar = 1 To Len(m_oFont.Name)
            .lfFaceName(lChar - 1) = CByte(Asc(Mid$(m_oFont.Name, lChar, 1)))
        Next lChar
        .lfHeight = -MulDiv(m_oFont.Size, GetDeviceCaps(lHdc, LOGPIXELSY), 72)
        .lfItalic = m_oFont.Italic
        .lfWeight = IIf(m_oFont.Bold, FW_BOLD, FW_NORMAL)
        .lfUnderline = m_oFont.Underline
        .lfStrikeOut = m_oFont.Strikethrough
        .lfCharSet = m_oFont.Charset
    End With
    DestroyFont
    m_lFont = CreateFontIndirect(uLF)
    SendMessageLong m_lLVHwnd, WM_SETFONT, m_lFont, True

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_Font", Err.Number)

End Property

'* Name           : DestroyFont
'* Purpose        : font cleanup
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Private Function DestroyFont() As Boolean

On Error GoTo Handler

    If m_lFont Then
        If DeleteObject(m_lFont) Then
            DestroyFont = True
            m_lFont = 0
        End If
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DestroyFont", Err.Number)
    
End Function

'* Name           : [get] p_ForeColor
'* Purpose        : retrieve list forecolor
'* Inputs         : none
'* Outputs        : ole color
'*********************************************
Public Property Get p_ForeColor() As OLE_COLOR

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    p_ForeColor = SendMessageLong(m_lLVHwnd, LVM_GETTEXTCOLOR, 0, 0)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ForeColor", Err.Number)

End Property

'* Name           : [let] p_ForeColor
'* Purpose        : change list forecolor
'* Inputs         : ole color
'* Outputs        : none
'*********************************************
Public Property Let p_ForeColor(ByVal PropVal As OLE_COLOR)

Dim lColor As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    OleTranslateColor PropVal, 0, lColor
    SendMessageLong m_lLVHwnd, LVM_SETTEXTCOLOR, 0, lColor
    ListRefresh

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ForeColor", Err.Number)

End Property

'* Name           : [get] p_FullRowSelect
'* Purpose        : retrieve full row select state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_FullRowSelect() As Boolean

    p_FullRowSelect = m_bFullRowSelect

End Property

'* Name           : [let] p_FullRowSelect
'* Purpose        : change full row select state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_FullRowSelect(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bFullRowSelect = PropVal
    If m_bFullRowSelect Then
        SetExtendedStyle LVS_EX_FULLROWSELECT, 0
    Else
        SetExtendedStyle 0, LVS_EX_FULLROWSELECT
    End If
    
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_FullRowSelect", Err.Number)

End Property

'* Name           : [get] p_GridLines
'* Purpose        : change gridlines state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_GridLines() As Boolean

    p_GridLines = m_bGridLines

End Property

'* Name           : [let] p_GridLines
'* Purpose        : change gridlines state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_GridLines(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bGridLines = PropVal
    If m_bGridLines Then
        SetExtendedStyle LVS_EX_GRIDLINES, 0
    Else
        SetExtendedStyle 0, LVS_EX_GRIDLINES
    End If
        
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_GridLines", Err.Number)

End Property

'* Name           : [get] p_HeaderDragDrop
'* Purpose        : retrieve drag and drop state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_HeaderDragDrop() As Boolean

    p_HeaderDragDrop = m_bDragDrop

End Property

'* Name           : [let] p_HeaderDragDrop
'* Purpose        : retrieve drag and drop state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Let p_HeaderDragDrop(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bDragDrop = PropVal
    If m_bDragDrop Then
        SetExtendedStyle LVS_EX_HEADERDRAGDROP, 0
    Else
        SetExtendedStyle 0, LVS_EX_HEADERDRAGDROP
    End If
        
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_HeaderDragDrop", Err.Number)

End Property

'* Name           : [get] p_HeaderFixedWidth
'* Purpose        : retrieve fixed width state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_HeaderFixedWidth() As Boolean

    p_HeaderFixedWidth = m_bHeaderFixed

End Property

'* Name           : [let] p_HeaderFixedWidth
'* Purpose        : change fixed width state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_HeaderFixedWidth(ByVal PropVal As Boolean)

    m_bHeaderFixed = PropVal

End Property

'* Name           : [get] p_HeaderFlat
'* Purpose        : change width state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Get p_HeaderFlat() As Boolean

    p_HeaderFlat = m_bHeaderFlat

End Property

'* Name           : [let] p_HeaderFlat
'* Purpose        : change width state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_HeaderFlat(ByVal PropVal As Boolean)

Dim lStyle      As Long
Dim lHwnd       As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHeaderFlat = PropVal
    lHwnd = HeaderHwnd()
    If lHwnd = 0 Then Exit Property
    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    If m_bHeaderFlat Then
        lStyle = lStyle And Not HDS_BUTTONS
    Else
        lStyle = lStyle Or HDS_BUTTONS
    End If
    SetWindowLong lHwnd, GWL_STYLE, lStyle

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_HeaderFlat", Err.Number)

End Property

'* Name           : [get] p_HeaderHide
'* Purpose        : retrieve header visible state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_HeaderHide() As Boolean

    p_HeaderHide = m_bHeaderHide

End Property

'* Name           : [let] p_HeaderHide
'* Purpose        : change header visible state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_HeaderHide(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHeaderHide = PropVal
    If m_bHeaderHide Then
        SetStyle LVS_NOCOLUMNHEADER, 0
    Else
        SetStyle 0, LVS_NOCOLUMNHEADER
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_HeaderHide", Err.Number)

End Property

'* Name           : [get] p_HideSelection
'* Purpose        : retrieve selection visible state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_HideSelection() As Boolean

    p_HideSelection = m_bHideSelection

End Property

'* Name           : [let] p_HideSelection
'* Purpose        : change selection visible state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_HideSelection(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bHideSelection = PropVal
    If m_bHideSelection Then
        SetStyle 0, LVS_SHOWSELALWAYS
    Else
        SetStyle LVS_SHOWSELALWAYS, 0
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_HideSelection", Err.Number)

End Property

'* Name           : [get] p_LabelTips
'* Purpose        : retrieve label tips state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_LabelTips() As Boolean

    p_LabelTips = m_bLabelTips

End Property

'* Name           : [let] p_LabelTips
'* Purpose        : change label tips state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_LabelTips(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bLabelTips = PropVal
    If m_bLabelTips Then
        SetExtendedStyle LVS_EX_LABELTIP, 0
    Else
        SetExtendedStyle 0, LVS_EX_LABELTIP
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_LabelTips", Err.Number)

End Property

'* Name           : [get] p_MultiSelect
'* Purpose        : retrieve multiselect state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_MultiSelect() As Boolean

    p_MultiSelect = m_bMultiSelect

End Property

'* Name           : [let] p_MultiSelect
'* Purpose        : change multiselect state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_MultiSelect(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bMultiSelect = PropVal
    If m_bMultiSelect Then
        SetStyle 0, LVS_SINGLESEL
    Else
        SetStyle LVS_SINGLESEL, 0
    End If
        
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_MultiSelect", Err.Number)

End Property

'* Name           : [get] p_OneClickActivate
'* Purpose        : retrieve oneclick state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_OneClickActivate() As Boolean

    p_OneClickActivate = m_bOneClick

End Property

'* Name           : [let] p_OneClickActivate
'* Purpose        : change oneclick state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_OneClickActivate(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bOneClick = PropVal
    If m_bOneClick Then
        SetExtendedStyle LVS_EX_ONECLICKACTIVATE, 0
    Else
        SetExtendedStyle 0, LVS_EX_ONECLICKACTIVATE
    End If
        
On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_OneClickActivate", Err.Number)

End Property

'* Name           : [get] p_ScrollBarFlat
'* Purpose        : retrieve scrollbar state
'* Inputs         : none
'* Outputs        : boolean
'*********************************************
Public Property Get p_ScrollBarFlat() As Boolean

    p_ScrollBarFlat = m_bScrollFlat

End Property

'* Name           : [let] p_ScrollBarFlat
'* Purpose        : change scrollbar state
'* Inputs         : boolean
'* Outputs        : none
'*********************************************
Public Property Let p_ScrollBarFlat(ByVal PropVal As Boolean)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_bScrollFlat = PropVal
    If m_bScrollFlat Then
        SetExtendedStyle LVS_EX_FLATSB, 0
    Else
        SetExtendedStyle 0, LVS_EX_FLATSB
    End If

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ScrollBarFlat", Err.Number)

End Property

'* Name           : [let] p_SelectedCount
'* Purpose        : retrieve selected count
'* Inputs         : none
'* Outputs        : long
'*********************************************
Public Property Get p_SelectedCount() As Long

    If m_lLVHwnd = 0 Then Exit Property
    p_SelectedCount = SendMessageLong(m_lLVHwnd, LVM_GETSELECTEDCOUNT, 0, 0)

End Property

'* Name           : [get] p_ViewMode
'* Purpose        : retrieve viewmode state
'* Inputs         : none
'* Outputs        : enum
'*********************************************
Public Property Get p_ViewMode() As eStyle

    p_ViewMode = m_ViewMode

End Property

'* Name           : [let] p_ViewMode
'* Purpose        : change viewmode state
'* Inputs         : enum
'* Outputs        : none
'*********************************************
Public Property Let p_ViewMode(ByVal PropVal As eStyle)

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Property
    m_ViewMode = PropVal
    SetStyle m_ViewMode, (LVS_ICON Or LVS_SMALLICON Or LVS_REPORT Or LVS_LIST)

On Error GoTo 0
Exit Property

Handler:
    RaiseEvent eHErrCond("p_ViewMode", Err.Number)

End Property


'* Name           : [let] p_ViewMode
'* Purpose        : change viewmode state
'* Inputs         : enum
'* Outputs        : none
'*********************************************
Public Function RemoveDuplicates() As Boolean

Dim lC          As Long
Dim lP          As Long
Dim lLb         As Long
Dim lUb         As Long
Dim cT          As Collection
Dim tHliTmp()   As HLIStc

On Error Resume Next

    ReDim tHliTmp(0)
    lLb = LBound(tHLIStc(0).Item)
    lUb = UBound(tHLIStc(0).Item)
    ReDim tHliTmp(0).Item(lLb To lUb)
    ReDim tHliTmp(0).lIcon(lLb To lUb)
    ReDim tHliTmp(0).SubItem1(lLb To lUb)
    ReDim tHliTmp(0).SubItem2(lLb To lUb)
    Set cT = New Collection
    lC = lUb
    lP = 0
    '/* only unique keys will be added
    Do
        cT.Add 1, tHLIStc(0).Item(lC)
        If Not Err.Number = 457 Then
            With tHliTmp(0)
                .Item(lP) = tHLIStc(0).Item(lC)
                .lIcon(lP) = tHLIStc(0).lIcon(lC)
                .SubItem1(lP) = tHLIStc(0).SubItem1(lC)
                .SubItem2(lP) = tHLIStc(0).SubItem2(lC)
                lP = lP + 1
            End With
        End If
        lC = lC - 1
    Loop While lC > 1
    
    With tHliTmp(0)
        ReDim Preserve .Item(lLb To lP)
        ReDim Preserve .lIcon(lLb To lP)
        ReDim Preserve .SubItem1(lLb To lP)
        ReDim Preserve .SubItem2(lLb To lP)
    End With
    
    DeAllocatePointer "a", True
    Erase tHLIStc
    ReDim tHLIStc(0)
    tHLIStc(0) = tHliTmp(0)
    Erase tHliTmp
    SetItemCount lP + 1
    
On Error GoTo 0

End Function

'* Name           : ItemsSort
'* Purpose        : sort items in the list
'* Inputs         : long
'* Outputs        : boolean
'*********************************************
Public Function ItemsSort(ByVal lColumn As Long, _
                          ByVal bDescending As Boolean) As Boolean

On Error GoTo Handler

    With tHLIStc(0)
        Select Case lColumn
        Case 0
            If bDescending Then
                SortControl .Item, 2
            Else
                SortControl .Item, 1
            End If
        Case 1
            If bDescending Then
                SortControl .SubItem1, 2
            Else
                SortControl .SubItem1, 1
            End If
        Case 2
            If bDescending Then
                SortControl .SubItem2, 2
            Else
                SortControl .SubItem2, 1
            End If
        End Select
        SetItemCount UBound(.Item) + 1
    End With

Handler:
    On Error GoTo 0

End Function

'>  Support
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


'* Name           : SetBorderStyle
'* Purpose        : change list borderstyle
'* Inputs         : handle, enum
'* Outputs        : none
'*********************************************
Private Sub SetBorderStyle(ByVal lHwnd As Long, _
                           ByVal eStyle As eBorderStyle)

On Error GoTo Handler

    Select Case eStyle
    Case [bLine]
        SetWindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        SetWindowStyle lHwnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    Case [bThin]
        SetWindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        SetWindowStyle lHwnd, GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    Case [bThick]
        SetWindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        SetWindowStyle lHwnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
    End Select

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetBorderStyle", Err.Number)

End Sub

'* Name           : SetBorderStyle
'* Purpose        : change list borderstyle
'* Inputs         : handle, enum
'* Outputs        : none
'*********************************************
Private Sub SetWindowStyle(ByVal lHwnd As Long, _
                           ByVal lType As Long, _
                           ByVal lStyle As Long, _
                           ByVal lStyleNot As Long)

Dim lNewStyle   As Long

On Error GoTo Handler

    lNewStyle = GetWindowLong(lHwnd, lType)
    lNewStyle = (lNewStyle And Not lStyleNot) Or lStyle
    SetWindowLong lHwnd, lType, lNewStyle
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetWindowStyle", Err.Number)
    
End Sub

'* Name           : SetStyle
'* Purpose        : change list style params
'* Inputs         : style, notstyle
'* Outputs        : none
'*********************************************
Private Sub SetStyle(ByVal lStyle As Long, _
                     ByVal lStyleNot As Long)

Dim lNewStyle   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Sub
    lNewStyle = GetWindowLong(m_lLVHwnd, GWL_STYLE)
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    SetWindowLong m_lLVHwnd, GWL_STYLE, lNewStyle
    SetWindowPos m_lLVHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
        
On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetStyle", Err.Number)

End Sub

'* Name           : SetExtendedStyle
'* Purpose        : change list extended style params
'* Inputs         : style, notstyle
'* Outputs        : none
'*********************************************
Private Sub SetExtendedStyle(ByVal lStyle As Long, _
                             ByVal lStyleNot As Long)


Dim lNewStyle   As Long

On Error GoTo Handler

    If m_lLVHwnd = 0 Then Exit Sub
    lNewStyle = SendMessageLong(m_lLVHwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lNewStyle = lNewStyle And Not lStyleNot
    lNewStyle = lNewStyle Or lStyle
    SendMessageLong m_lLVHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lNewStyle
        
On Error GoTo 0
Exit Sub

Handler:
    RaiseEvent eHErrCond("SetExtendedStyle", Err.Number)

End Sub

Public Sub ListRefresh()

    InvalidateRect m_lLVHwnd, ByVal 0&, 0&
    UpdateWindow m_lLVHwnd

End Sub

Public Sub SetItemCount(ByVal nItems As Long)

    SendMessage m_lLVHwnd, LVM_SETITEMCOUNT, nItems, LVSICF_NOINVALIDATEALL
    m_lItemsCnt = nItems
    InitCheckBoxes m_lItemsCnt

End Sub

Private Function ItemChecked(lIndex As Long) As Boolean

On Error GoTo Handler

    ItemChecked = m_lCheckState(lIndex) > 0

Handler:
    On Error GoTo 0
    
End Function

Private Function CheckToggle(ByVal lItem As Long) As Boolean

On Error GoTo Handler

    If ItemChecked(lItem) Then
        m_lCheckState(lItem) = 0
        CheckToggle = 0
    Else
        m_lCheckState(lItem) = 1
        CheckToggle = 1
    End If

Handler:
    On Error GoTo 0

End Function

Private Function InitCheckBoxes(ByVal lCount As Long)

    ReDim m_lCheckState(lCount)
    m_bCheckInit = True
    
End Function

'* Name           : ColumnReorder
'* Purpose        : reorder columns to accomodate checkbox
'* Inputs         : none
'* Outputs        : none
'*********************************************
Public Function ColumnReorder(ByVal bRemCheckbox As Boolean)

Dim aText()     As String
Dim aWidth()    As Long
Dim lCt         As Long
Dim lUct        As Long

On Error Resume Next

    lUct = p_ColumnCount
    ReDim aText(lUct)
    ReDim aWidth(lUct)
    
    For lCt = 0 To lUct
        aText(lCt) = Trim$(p_ColumnText(lCt))
        aWidth(lCt) = p_ColumnWidth(lCt)
    Next lCt
    
    For lCt = lUct To 0 Step -1
        ColumnRemove lCt
    Next lCt
    
    If bRemCheckbox Then
        For lCt = 0 To lUct
            ColumnAdd lCt + 1, aText(lCt), aWidth(lCt)
        Next lCt
    Else
        For lCt = 1 To lUct
            ColumnAdd lCt - 1, aText(lCt - 1), aWidth(lCt - 1)
        Next lCt
    End If

On Error GoTo 0

End Function

Private Function DestroyList() As Boolean

    DestroyImlHeader
    DestroyImlSmall
    DestroyImlLarge
    If m_lLVHwnd Then
        If DestroyWindow(m_lLVHwnd) Then
            DestroyList = True
            m_lLVHwnd = 0
        End If
    End If

End Function

Private Function LOWORD(ByVal dwValue As Long) As Long
' Returns the low 16-bit Long from a 32-bit long Long

    CopyMemory LOWORD, dwValue, 2&

End Function

Private Function HIWORD(ByVal nValue As Long) As Long
' returns the high 16-bit Long from a 32-bit long Long

    CopyMemory HIWORD, ByVal VarPtr(nValue) + 2, 2&

End Function

Private Sub List_Attatch(ByVal lHwnd As Long)
'/* attatch messages

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Attach_Message Me, m_lParentHwnd, WM_NOTIFY
        .Attach_Message Me, m_lParentHwnd, WM_SETFOCUS
        .Attach_Message Me, lHwnd, WM_TIMER
        .Attach_Message Me, lHwnd, WM_LBUTTONDOWN
        .Attach_Message Me, lHwnd, WM_MBUTTONDOWN
        .Attach_Message Me, lHwnd, WM_RBUTTONDOWN
        .Attach_Message Me, lHwnd, WM_NOTIFY
        .Attach_Message Me, lHwnd, WM_SETFOCUS
        .Attach_Message Me, lHwnd, WM_KEYDOWN
        .Attach_Message Me, lHwnd, WM_CHAR
        .Attach_Message Me, lHwnd, WM_KEYUP
        .Attach_Message Me, lHwnd, WM_MOUSEMOVE
        .Attach_Message Me, lHwnd, WM_LBUTTONUP
        .Attach_Message Me, lHwnd, WM_RBUTTONUP
        .Attach_Message Me, lHwnd, WM_MBUTTONUP
    End With

End Sub

Private Sub List_Detatch(ByVal lHwnd As Long)
'/* detatch messages

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Detach_Message Me, m_lParentHwnd, WM_NOTIFY
        .Detach_Message Me, m_lParentHwnd, WM_SETFOCUS
        .Detach_Message Me, lHwnd, WM_TIMER
        .Detach_Message Me, lHwnd, WM_LBUTTONDOWN
        .Detach_Message Me, lHwnd, WM_MBUTTONDOWN
        .Detach_Message Me, lHwnd, WM_RBUTTONDOWN
        .Detach_Message Me, lHwnd, WM_NOTIFY
        .Detach_Message Me, lHwnd, WM_SETFOCUS
        .Detach_Message Me, lHwnd, WM_KEYDOWN
        .Detach_Message Me, lHwnd, WM_CHAR
        .Detach_Message Me, lHwnd, WM_KEYUP
        .Detach_Message Me, lHwnd, WM_MOUSEMOVE
        .Detach_Message Me, lHwnd, WM_LBUTTONUP
        .Detach_Message Me, lHwnd, WM_RBUTTONUP
        .Detach_Message Me, lHwnd, WM_MBUTTONUP
    End With

End Sub

Private Property Get MISubclass_MsgResponse() As EMsgResponse
'/* message status property

    MISubclass_MsgResponse = emrPreProcess

End Property

Private Property Let MISubclass_MsgResponse(ByVal RHS As EMsgResponse)

    '<STUB>

End Property

Private Function MISubclass_WindowProc(ByVal lHwnd As Long, _
                                       ByVal iMsg As Long, _
                                       ByVal wParam As Long, _
                                       ByVal lParam As Long) As Long

Dim hIcon                   As Long
Dim lCode                   As Long
Dim lHandle                 As Long
Dim lItem                   As Long
Dim sTemp                   As String
Dim tNmhdr                  As NMHDR
Dim tNmList                 As NM_LISTVIEW
Dim tDisp                   As LV_DISPINFO
Dim LVI                     As LV_ITEM
Dim tGRedraw                As NMLVCUSTOMDRAW

    '/* a little focus hack, how easy was that? ;o)
    If lHwnd = m_lParentHwnd Then
        If iMsg = WM_SETFOCUS Then
            SetFocus m_lLVHwnd
        End If
    End If
    
    Select Case iMsg
    Case WM_NOTIFY
        CopyMemory tNmhdr, ByVal lParam, LenB(tNmhdr)
        With tNmhdr
            If .code = TTN_GETDISPINFO Then
                MISubclass_WindowProc = 1
                Exit Function
            End If
            lCode = .code
            lHandle = .hwndFrom
        End With
        
        Select Case lHandle
        '/* static header
        Case m_lHdrHwnd
            If lCode = HDN_ITEMCHANGING Then
                If m_bHeaderFixed Then
                    MISubclass_WindowProc = 1
                End If
            End If
            
            '/* custom header colors
            If lCode = NM_CUSTOMDRAW Then
                If Not m_bCustomHeader Then Exit Function
                CopyMemory tGRedraw, ByVal lParam, Len(tGRedraw)
                If tGRedraw.nmcmd.dwDrawStage = CDDS_PREPAINT Then
                    MISubclass_WindowProc = CDRF_NOTIFYITEMDRAW
                    Exit Function
                End If
            
                If tGRedraw.nmcmd.dwDrawStage = CDDS_ITEMPREPAINT Then
                    SetTextColor tGRedraw.nmcmd.hdc, m_oHdrForeClr
                    SetBkColor tGRedraw.nmcmd.hdc, m_oHdrBkClr
                Else
                    MISubclass_WindowProc = CDRF_DODEFAULT
                    Exit Function
                End If
        
                If tGRedraw.nmcmd.dwDrawStage = CDDS_ITEMPOSTPAINT Then
                    MISubclass_WindowProc = CDRF_DODEFAULT
                    Exit Function
                End If
            End If
            
        Case m_lLVHwnd
            Select Case lCode
            Case LVN_COLUMNCLICK, LVN_ITEMCHANGING
                CopyMemory tNmList, ByVal lParam, LenB(tNmList)
            Case LVN_BEGINLABELEDIT, LVN_ENDLABELEDIT, LVN_GETDISPINFO
                CopyMemory tDisp, ByVal lParam, LenB(tDisp)
            End Select
                
            Select Case lCode
            '/* column click
            Case LVN_COLUMNCLICK
                If Not m_bUseSorted Then Exit Function
                Dim bDesc As Boolean
                RaiseEvent eHColumnClick(tNmList.iSubItem)
                '/* swap sort icon
                p_ColumnIcon(0) = -1
                p_ColumnIcon(1) = -1
                p_ColumnIcon(2) = -1
                If p_ColumnIcon(tNmList.iSubItem) = 0 Then
                    p_ColumnIcon(tNmList.iSubItem) = 1
                    bDesc = True
                Else
                    p_ColumnIcon(tNmList.iSubItem) = 0
                End If
                '/* column and sort direction
                With tHLIStc(0)
                    Select Case tNmList.iSubItem
                    Case 0
                        If bDesc Then
                            SortControl .Item, 1
                        Else
                            SortControl .Item, 2
                        End If
                    Case 1
                        If bDesc Then
                            SortControl .SubItem1, 1
                        Else
                            SortControl .SubItem1, 2
                        End If
                    Case 2
                        If bDesc Then
                            SortControl .SubItem2, 1
                        Else
                            SortControl .SubItem2, 2
                        End If
                    End Select
                End With
                
            '/* item changed
            Case LVN_ITEMCHANGED
                CopyMemory tNmList, ByVal lParam, Len(tNmList)
                With tNmList
                    If .uOldState Then
                        If ((.uNewState And LVIS_STATEIMAGEMASK) <> (.uOldState And LVIS_STATEIMAGEMASK)) Then
                            RaiseEvent eHItemCheck(.iItem)
                        End If
                    Else
                        If Not m_bFirstItem Then
                            If ((.uNewState And LVIS_SELECTED)) Then
                                RaiseEvent eHItemClick(.iItem)
                            End If
                        End If
                    End If
                End With

            '/* checkbox click
            Case LVIS_CHKCLICK
                If tDisp.Item.iSubItem = 0 And m_bCheckBoxes Then
                    CopyMemory tNmList, ByVal lParam, Len(tNmList)
                    CheckToggle tNmList.iItem
                    ListRefresh
                End If
                    
            '/* list change callback
            Case LVN_GETDISPINFO
                '/* indirect callback method:
                '/* use the item callback as link to an external database
                '/* internal method:
                '/* fetch item by index
                With tDisp.Item
                    If m_bIndirectMode Then
                        RaiseEvent eHIndirect(.iItem, .iSubItem, .mask, sTemp, hIcon)
                    Else
                        'p_ItemIcon(.iItem) = tHLIStc(0).lIcon(.iItem)
                        '/* display items by qualified pointer index
                        '/* this will probably be revised..
                        If m_bUseSorted And m_bSorted Then
                            Select Case .iSubItem
                            Case 0
                                sTemp = tHLIStc(0).Item(m_lPtr(.iItem))
                            Case 1
                                sTemp = tHLIStc(0).SubItem1(m_lPtr(.iItem))
                            Case 2
                                sTemp = tHLIStc(0).SubItem2(m_lPtr(.iItem))
                            End Select
                        '/* normal retrieval
                        Else
                            Select Case .iSubItem
                            Case 0
                                sTemp = tHLIStc(0).Item(.iItem)
                            Case 1
                                sTemp = tHLIStc(0).SubItem1(.iItem)
                            Case 2
                                sTemp = tHLIStc(0).SubItem2(.iItem)
                            End Select
                        End If
                    End If
                End With
            
                With tDisp.Item
                    If .iSubItem = 0 Then
                        If .mask And LVIF_TEXT Then
                            '/ copy text
                            If Len(sTemp) > .cchTextMax Then
                                sTemp = left$(sTemp, .cchTextMax)
                            End If
                            lstrcpyToPointer .pszText, sTemp
                        End If
                        If .mask And LVIF_STATE Then
                            Debug.Print "check"
                        End If
                        If .mask And LVIF_IMAGE Then
                            If m_bCheckBoxes Then
                                Select Case ItemChecked(.iItem)
                                Case 0
                                    .State = LVIS_UNCHECKED
                                Case 1
                                    .State = LVIS_CHECKED
                                End Select
                            End If
                            .iImage = tHLIStc(0).lIcon(tDisp.Item.iItem)
                            .mask = LVIF_IMAGE Or LVIF_TEXT Or LVIF_STATE           '<-why didn't I see this sooner?, doh!
                            .stateMask = LVIS_OVERLAYMASK Or LVIS_STATEIMAGEMASK    '<- doesn't work right in small icon though, no idea why..
                        End If
                            
                        If .mask And LVIF_INDENT Then
                            .iIndent = 0
                        End If
                    Else
                        If .mask And LVIF_TEXT Then
                            '/ copy text
                            If Len(sTemp) > .cchTextMax Then
                                sTemp = left$(sTemp, .cchTextMax)
                            End If
                            lstrcpyToPointer .pszText, sTemp
                        End If
                        If m_bSubItemImage Then
                            .iImage = tHLIStc(0).lIcon(tDisp.Item.iItem)
                            .mask = LVIF_IMAGE Or LVIF_TEXT
                            .stateMask = LVIS_OVERLAYMASK
                        End If
                    End If
                End With
                CopyMemory ByVal lParam, tDisp, LenB(tDisp)
            End Select
        End Select
    End Select

End Function

'**********************************************************************
'*                              STORAGE
'**********************************************************************

'                   <<+>>
'||<<<<<<<<<<<<<<<< Build >>>>>>>>>>>>>>>>||
'                   <<+>>

Private Function ArrayCheck(ByRef vArray As Variant) As Boolean
'/* validity test

On Error GoTo Handler

    '/* an array
    If IsArray(vArray) Then
        On Error Resume Next
        '/* dimensioned
        If IsError(UBound(vArray)) Then
            GoTo Handler
        End If
        On Error GoTo 0
    Else
        GoTo Handler
    End If

    ArrayCheck = True

On Error GoTo 0
Exit Function

Handler:

End Function

Private Function ArrayExists(ByRef vArray As Variant) As Boolean

On Error Resume Next

    If IsError(UBound(vArray)) Then
        GoTo Handler
    ElseIf UBound(vArray) < 2 Then
        GoTo Handler
    End If

    '/* success
    ArrayExists = True

Handler:
    On Error GoTo 0

End Function

Public Function ArrayRemLong(ByRef lArray() As Long, _
                             ByVal lPos As Long) As Boolean
Dim lLb     As Long
Dim lUb     As Long

On Error GoTo Handler

    lLb = LBound(lArray)
    lUb = UBound(lArray)

    If Not ArrayCheck(lArray) Then Exit Function
    If lPos > lUb Then lPos = lUb
    If lPos < lLb Then lPos = lLb
    If lPos = lUb Then
        ReDim Preserve lArray(lUb - 1)
        Exit Function
    End If

    CopyMemory lArray(lPos), lArray(lPos + 1), (lUb - lLb - lPos) * Len(lArray(lPos))
    ReDim Preserve lArray(lUb - 1)
    '/* success
    ArrayRemLong = True

Handler:
    On Error GoTo 0

End Function

Public Function ArrayRemString(ByRef sArray() As String, _
                               ByVal lPos As Long) As Boolean

Dim lLb     As Long
Dim lUb     As Long
Dim lPtr    As Long

On Error GoTo Handler

    If Not ArrayCheck(sArray) Then Exit Function
    lLb = LBound(sArray)
    lUb = UBound(sArray)
    If lPos > lUb Then lPos = lUb
    If lPos < lLb Then lPos = lLb
    If lPos = lUb Then
        ReDim Preserve sArray(lUb - 1)
        Exit Function
    End If
    
    lPtr = StrPtr(sArray(lPos))
    CopyMemory ByVal VarPtr(sArray(lPos)), ByVal VarPtr(sArray(lPos + 1)), (lUb - lPos) * 4
    CopyMemory ByVal VarPtr(sArray(lUb)), lPtr, 4
    ReDim Preserve sArray(lUb - 1)
    '/* success
    ArrayRemString = True

Handler:
    On Error GoTo 0

End Function

Private Function DeAllocatePointer(ByVal sKey As String, _
                                   Optional ByVal bPurge As Boolean) As Boolean

'/* resolve or purge memory pointers

Dim lPtr    As Long
Dim lC      As Long

On Error GoTo Handler

    If Not bPurge Then
        '/* get the pointer
        lPtr = c_PtrMem.Item(sKey)
        If lPtr = 0 Then GoTo Handler
        '/* release the memory
        CopyMemBr ByVal lPtr, 0&, &H4
    Else
        '/* destroy the struct last
        For lC = c_PtrMem.Count To 1 Step -1
            If Not CLng(c_PtrMem.Item(lC)) = 0 Then
                lPtr = CLng(c_PtrMem.Item(lC))
                CopyMemBr ByVal lPtr, 0&, &H4
            End If
        Next lC
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("DeAllocatePointer", Err.Number)

End Function


'* Name           : Sort_Control
'* Purpose        : sorting hub
'* Inputs         : sort type -enum
'* Outputs        : boolean
'*********************************************
Public Function SortControl(ByRef sArray() As String, _
                            ByVal lSortType As Long) As Boolean

On Error GoTo Handler

'1GHz celeron, 100fsb, 256 mb/r
'100,000 * len- 8, semi sorted, case sensitive
'Idx TriQuickSort
'avg: 1.59
'Idx Qsort
'avg: 1.26

    If Not ArrayCheck(sArray) Then GoTo Handler
    '/* array less then min dimensions
    If Not ArrayExists(sArray) Then GoTo Handler
    '/* default sort
    If lSortType = 0 Then lSortType = 1
    '/* load a new pointer index
    QSIInitPtr LBound(sArray), UBound(sArray), m_lPtr
    
    '/* Case - lCp
    '/* &h1 no case, &h0 case(binary)
    '/* Order - lDir
    '/* &h1 ascend, &hffff descend (+1 more, -1 less)
    Select Case lSortType
    '/* ascending case sensitive
    Case 1
        QSISort sArray, m_lPtr, LBound(sArray), UBound(sArray), &H0, &H1
    '/* reverse case sensitive
    Case 2
        QSISort sArray, m_lPtr, LBound(sArray), UBound(sArray), &H0, &HFFFF
    '/* forward case insensitive
    Case 3
        QSISort sArray, m_lPtr, LBound(sArray), UBound(sArray), &H1, &H1
    '/* reverse case insensitive
    Case 4
        QSISort sArray, m_lPtr, LBound(sArray), UBound(sArray), &H1, &HFFFF
    End Select
    
    'TestSort
    SetItemCount UBound(sArray) + 1
    '/* success
    m_bSorted = True
    SortControl = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eHErrCond("Sort_Control", Err.Number)

End Function

Private Sub QSISort(sA() As String, _
                    lIdxA() As Long, _
                    ByVal lbA As Long, _
                    ByVal ubA As Long, _
                    ByVal lCp As Long, _
                    ByVal lDr As Long)

'/* based on the awesome indexed sort by Rde (Rohan) w/ mods

Dim lo          As Long
Dim hi          As Long
Dim cnt         As Long
Dim Item        As String
Dim lpStr       As Long
Dim idxItem     As Long
Dim lpS         As Long

    '/* pre execution check
    If Not UBound(sA) > 0 Then Exit Sub
    '/* Allow for worst case senario + some
    hi = ((ubA - lbA) \ m8) + m32
    '/* Stack to hold pending lower boundries
    ReDim lbs(m1 To hi) As Long
    '/* Stack to hold pending upper boundries
    ReDim ubs(m1 To hi) As Long
    '/* Cache pointer to the string variable                                                '*** Change History ***
    lpStr = GetVarPtr(Item)                                                                 '<- Change 1. VarPtr exchanged for direct access
    '/* Cache pointer to the string array                                                   '<- pointer call in typelib, -GetVarPtr, avoiding runtime
    lpS = GetVarPtr(sA(lbA)) - (lbA * m4)                                                   '<- error checking, and saving a call per execution
                                                                                            '<- this is a considerable gain in a busy loop structure
    '/* Get pivot index position                                                            '<- warning- do not pass a null pointer though!
    Do: hi = ((ubA - lbA) \ m2) + lbA
        '/* Grab current value into item
        CopyMemBv lpStr, lpS + (lIdxA(hi) * m4), m4                                         '<- Change 2. using faster typelib copymembv,
        '/* Grab current index                                                              '<- using typelib bypasses runtime getlasterror
        idxItem = lIdxA(hi): lIdxA(hi) = lIdxA(ubA)                                         '<- check, saving a secondary call on each execution.
        '/* Set bounds                                                                      '<- function call pointer is also resolved directly to
        lo = lbA: hi = ubA                                                                  '<- routine rather then at runtime when using 'declare'
        '/* Storm right in
        Do
            If Not StrComp(Item, sA(lIdxA(lo)), lCp) = lDr Then                             '<- Change 3. simplified not structure means 2 less
                lIdxA(hi) = lIdxA(lo)                                                       '<- instructions per iteration, and 3%+ performance gain
                hi = hi - m1
                Do
                    If Not StrComp(sA(lIdxA(hi)), Item, lCp) = lDr Then
                        lIdxA(lo) = lIdxA(hi)
                        Exit Do
                    End If
                    hi = hi - m1
                Loop Until hi = lo
                '/* Found swaps or out of loop
                If hi = lo Then Exit Do
            End If
            lo = lo + m1
        Loop While hi > lo                                                                  '<- Change 4. Do While/Loop changed to faster Do/Loop While,
        '/* Re-assign current                                                               '<- assembles without additional jmp instruction
        lIdxA(hi) = idxItem
        If (lbA < lo - m1) Then
            If (ubA > lo + m1) Then cnt = cnt + m1: lbs(cnt) = lo + m1: ubs(cnt) = ubA
            ubA = lo - m1
        ElseIf (ubA > lo + m1) Then
            lbA = lo + m1
        Else
            If cnt = m0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - m1
        End If
    Loop: CopyMem ByVal lpStr, 0&, m4
    
End Sub

Private Sub QSIInitPtr(ByVal lLb As Long, _
                       ByVal lUb As Long, _
                       ByRef aPtr() As Long)

'/* initialize the pointer array
Dim lC As Long

    Erase aPtr
    ReDim aPtr(lLb To lUb)
    lC = lLb
    
    Do
        aPtr(lC) = lC
        lC = lC + 1
    Loop Until lC = lUb
    
End Sub

Private Sub TestSort(ByRef sArray() As String)

Dim lC As Long

    For lC = LBound(sArray) To UBound(sArray) Step 1000
        'Debug.Print sArray(lC)
        Debug.Print sArray(m_lPtr(lC))
    Next lC
    
End Sub

Private Sub Class_Terminate()

    List_Detatch m_lLVHwnd
    Set m_GSubclass = Nothing
    DestroyList
    DeAllocatePointer "a", True

End Sub
