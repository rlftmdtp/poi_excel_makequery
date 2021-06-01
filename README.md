# poi_excel_makequery
엑셀의 수만개의 데이터를 DB에 입력하기 위해 엑셀의 데이터를 읽은 후 Insert문을 생성해주는 소스

회사프로젝트 자바 1.7사용 그에 호환되는 POI 3.7, XMLBeans 2.6 버전사용


XSSFWorkBook는 .xlsx 확장자 엑셀 읽기전용

(.xls는 HSSFWorkbook를 사용해야함)

-------------------------------------------------------------------------------------------
진행하면서 해결했던 에러

1.org.apache.commons.collections4.listvaluedmap  에러  -> commons-collections4-x.x.jar 추가

2.org.apache.xmlbeans.xmlexception 에러 -> XMLBeans-2.6.0 추가
(http://jar.fyicenter.com/2945_Donwload_Apache_XMLBeans-2_6_0_zip.html)
