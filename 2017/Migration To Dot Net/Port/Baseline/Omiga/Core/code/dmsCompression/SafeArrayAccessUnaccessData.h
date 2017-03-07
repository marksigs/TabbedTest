#ifndef __SAFEARRAYACCESSUNACCESSDATA_H
#define __SAFEARRAYACCESSUNACCESSDATA_H


class CSafeArrayAccessUnaccessData
{
public:
	CSafeArrayAccessUnaccessData() : m_psa(NULL) {}
	HRESULT Access(SAFEARRAY* psa, void HUGEP** ppvData) {m_psa = psa; return SafeArrayAccessData(m_psa, ppvData);}
	~CSafeArrayAccessUnaccessData() {if (m_psa != NULL) SafeArrayUnaccessData(m_psa);}

private:
	SAFEARRAY* m_psa;
};


#endif //__SAFEARRAYACCESSUNACCESSDATA_H
