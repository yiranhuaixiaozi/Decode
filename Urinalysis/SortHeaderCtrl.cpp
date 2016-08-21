#include "stdafx.h"
#include "SortHeaderCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CSortHeaderCtrl::CSortHeaderCtrl()
	: m_iSortColumn( -1 )
	, m_bSortAscending( TRUE )
{
}

CSortHeaderCtrl::~CSortHeaderCtrl()
{
}


BEGIN_MESSAGE_MAP(CSortHeaderCtrl, CHeaderCtrl)
	//{{AFX_MSG_MAP(CSortHeaderCtrl)
		ON_WM_PAINT()

		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSortHeaderCtrl message handlers

void CSortHeaderCtrl::SetSortArrow( const int iSortColumn, const BOOL bSortAscending )
{
	m_iSortColumn = iSortColumn;
	m_bSortAscending = bSortAscending;

	// change the item to owner drawn.
	HD_ITEM hditem;

	hditem.mask = HDI_FORMAT;
	VERIFY( GetItem( iSortColumn, &hditem ) );
	hditem.fmt |= HDF_OWNERDRAW;
	VERIFY( SetItem( iSortColumn, &hditem ) );

	// invalidate the header control so it gets redrawn
	t_ColorText = RGB(0,0,0);
	t_ColorBk  = RGB(215,233,245);
	Invalidate();
}


void CSortHeaderCtrl::DrawItem( LPDRAWITEMSTRUCT lpDrawItemStruct )
{
	// attath to the device context.
	CDC dc;
	VERIFY( dc.Attach( lpDrawItemStruct->hDC ) );

	// save the device context.
	const int iSavedDC = dc.SaveDC();

	// get the column rect.
	CRect rc( lpDrawItemStruct->rcItem );

	// set the clipping region to limit drawing within the column.
/*	CRgn rgn;
	VERIFY( rgn.CreateRectRgnIndirect( &rc ) );
	(void)dc.SelectObject( &rgn );
	VERIFY( rgn.DeleteObject() );
*/
//	CRgn rgn;
//	rgn.CreateRectRgnIndirect( &rc ) ;
//	(void)dc.SelectObject( &rgn );
//	rgn.DeleteObject();

	CBrush Br3;
	CBrush *ob;
	Br3.CreateSolidBrush(GetSysColor( COLOR_3DFACE ));
	ob = dc.SelectObject(&Br3);
	dc.FillRect( rc, &Br3);
	dc.SelectObject(ob);
	DeleteObject(Br3);

	// draw the background,

	// get the column text and format.
	TCHAR szText[ 256 ];
	HD_ITEM hditem;

	hditem.mask = HDI_TEXT | HDI_FORMAT;
	hditem.pszText = szText;
	hditem.cchTextMax = 255;

	VERIFY( GetItem( lpDrawItemStruct->itemID, &hditem ) );

	// determine the format for drawing the column label.
	UINT uFormat = DT_SINGLELINE | DT_NOPREFIX | DT_NOCLIP | DT_VCENTER | DT_END_ELLIPSIS ;

	if( hditem.fmt & HDF_CENTER)
		uFormat |= DT_CENTER;
	else if( hditem.fmt & HDF_RIGHT)
		uFormat |= DT_RIGHT;
	else
		uFormat |= DT_LEFT;

	// adjust the rect if the mouse button is pressed on it.
	if( lpDrawItemStruct->itemState == ODS_SELECTED )
	{
		rc.left++;
		rc.top += 2;
		rc.right++;
	}

	CRect rcIcon( lpDrawItemStruct->rcItem );
	const int iOffset = ( rcIcon.bottom - rcIcon.top ) / 4;

	// adjust the rect further if the sort arrow is to be displayed.
	if( lpDrawItemStruct->itemID == (UINT)m_iSortColumn )
		rc.right -= 3 * iOffset;

	rc.left += iOffset;
	rc.right -= iOffset;

	// draw the column label.
	if( rc.left < rc.right )
		(void)dc.DrawText( szText, -1, rc, uFormat );

	// draw the sort arrow.
	if( lpDrawItemStruct->itemID == (UINT)m_iSortColumn )
	{
		// set up the pens to use for drawing the arrow.
		CPen penLight;
		CPen* pOldPen;

		if( m_bSortAscending )
		{
			// draw the arrow pointing upwards.
			penLight.CreatePen( PS_SOLID, 1, GetSysColor( COLOR_3DHILIGHT ));
			pOldPen = dc.SelectObject(&penLight);
			dc.MoveTo( rcIcon.right - 2 * iOffset, iOffset);
			dc.LineTo( rcIcon.right - iOffset, rcIcon.bottom - iOffset - 1 );
			dc.LineTo( rcIcon.right - 3 * iOffset - 2, rcIcon.bottom - iOffset - 1 );
			dc.SelectObject(pOldPen);
			DeleteObject(penLight);

//			(void)dc.SelectObject( &penShadow );

			penLight.CreatePen( PS_SOLID, 1, GetSysColor( COLOR_3DSHADOW ) );
			pOldPen = dc.SelectObject(&penLight);
			dc.MoveTo( rcIcon.right - 3 * iOffset - 1, rcIcon.bottom - iOffset - 1 );
			dc.LineTo( rcIcon.right - 2 * iOffset, iOffset - 1);		
			dc.SelectObject(pOldPen);
			DeleteObject(penLight);
		}
		else
		{
			// draw the arrow pointing downwards.
			penLight.CreatePen( PS_SOLID, 1, GetSysColor( COLOR_3DHILIGHT ));
			pOldPen = dc.SelectObject(&penLight);
			dc.MoveTo( rcIcon.right - iOffset - 1, iOffset );
			dc.LineTo( rcIcon.right - 2 * iOffset - 1, rcIcon.bottom - iOffset );
			dc.SelectObject(pOldPen);
			DeleteObject(penLight);

//			(void)dc.SelectObject( &penShadow );
			penLight.CreatePen( PS_SOLID, 1, GetSysColor( COLOR_3DSHADOW ) );
			pOldPen = dc.SelectObject(&penLight);
			dc.MoveTo( rcIcon.right - 2 * iOffset - 2, rcIcon.bottom - iOffset );
			dc.LineTo( rcIcon.right - 3 * iOffset - 1, iOffset );
			dc.LineTo( rcIcon.right - iOffset - 1, iOffset );		
			dc.SelectObject(pOldPen);
			DeleteObject(penLight);
		}

		// restore the pen.
//		(void)dc.SelectObject( pOldPen );
	}

	// restore the previous device context.
	VERIFY( dc.RestoreDC( iSavedDC ) );

	// detach the device context before returning.
	(void)dc.Detach();
}


void CSortHeaderCtrl::Serialize( CArchive& ar )
{
	if( ar.IsStoring() )
	{
		const int iItemCount = GetItemCount();
		if( iItemCount != -1 )
		{
			ar << iItemCount;

			HD_ITEM hdItem = { 0 };
			hdItem.mask = HDI_WIDTH;

			for( int i = 0; i < iItemCount; i++ )
			{
				VERIFY( GetItem( i, &hdItem ) );
				ar << hdItem.cxy;
			}
		}
	}
	else
	{
		int iItemCount;
		ar >> iItemCount;
		
		if( GetItemCount() != iItemCount )
			TRACE0( _T("Different number of columns in registry.") );
		else
		{
			HD_ITEM hdItem = { 0 };
			hdItem.mask = HDI_WIDTH;

			for( int i = 0; i < iItemCount; i++ )
			{
				ar >> hdItem.cxy;
				VERIFY( SetItem( i, &hdItem ) );
			}
		}
	}
}


void CSortHeaderCtrl::OnPaint() 
{
	 CPaintDC dc(this); // device context for painting
 
	 // TODO: Add your message handler code here
	CRect rect;
	GetClientRect(rect);
	dc.FillSolidRect(rect,t_ColorBk);   //重绘标题栏颜色
	int nItems = GetItemCount();
	CRect rectItem;
	CPen m_pen(PS_SOLID,1,RGB(211,211,211));      //分隔线颜色
	CPen * pOldPen=dc.SelectObject(&m_pen);

	CFont ftPrint;
	
	CSize szFtPrint;
	int iFtPrint= 0;	// fonts sizes
	SetFont(-16,0,0,0,400,0,0,0,0,3,2,1,34);
	ftPrint.CreateFontIndirect(&m_Logfont);
//	CString cs = "6";							
	CFont * of = dc.SelectObject(&ftPrint);//得到字体尺寸
//	szFtPrint = dc.GetTextExtent(cs);
//	dc.GetCurrentFont()->GetLogFont(&m_Logfont);
//	strcpy(m_Logfont.lfFaceName, "Arial");
//	m_Logfont.lfCharSet=134;
//	m_Logfont.lfHeight=-18;
//	m_Logfont.lfWidth=0;
//	ftPrint.CreateFontIndirect(&m_Logfont);
//	CFont * of = dc.SelectObject(&ftPrint);
	dc.SetTextColor(t_ColorText);     //字体颜色

	for(int i = 0; i <nItems; i++)                    //对标题的每个列进行重绘
	{  
		GetItemRect(i, &rectItem);
		rectItem.top+=2;
		rectItem.bottom+=2; 
		dc.MoveTo(rectItem.right,rect.top);                //重绘分隔栏
		dc.LineTo(rectItem.right,rectItem.bottom);
		TCHAR buf[256];
		HD_ITEM hditem;
 		hditem.mask = HDI_TEXT | HDI_FORMAT | HDI_ORDER;
		hditem.pszText = buf;
		hditem.cchTextMax = 255;
		GetItem( i, &hditem );                                       //获取当然列的文字
		UINT uFormat = DT_SINGLELINE | DT_NOPREFIX | DT_TOP |DT_CENTER | DT_END_ELLIPSIS|DT_VCENTER  ;
		dc.DrawText(buf, &rectItem, uFormat);           //重绘标题栏的文字
	 }
	 dc.SelectObject(of);
	 DeleteObject(ftPrint);
	 dc.SelectObject(pOldPen);
}


void CSortHeaderCtrl::SetFont(int A1,int A2,int A3,int A4,int A5,int A6,int A7,int A8,int A9,int A10,int A11,int A12,int A13)
{
	m_Logfont.lfHeight =A1;
	m_Logfont.lfWidth = A2;
	m_Logfont.lfEscapement = A3;
	m_Logfont.lfOrientation = A4;
	m_Logfont.lfWeight = A5;
	m_Logfont.lfItalic = A6;
	m_Logfont.lfUnderline = A7;
	m_Logfont.lfStrikeOut = A8;
	m_Logfont.lfCharSet = A9;
	m_Logfont.lfOutPrecision =A10;
	m_Logfont.lfClipPrecision = A11;
	m_Logfont.lfQuality = A12;
	m_Logfont.lfPitchAndFamily = A13;
	strcpy(m_Logfont.lfFaceName, "Arial");        // request a face name "Arial"

}
//	SetFont(-17,0,0,0,500,0,0,0,0,3,2,1,34);
