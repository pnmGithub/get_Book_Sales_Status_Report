## [도서판매현황 리포트 작성]  ##
> 압축된 도서정보 파일을 압축해제하여 분야별 도서판매현황 리포트 작성
> 제공한 템플릿 양식을 Copy하여 작업, 결과파일 작성    
> REFramework로 구현 (QueueItem이 아닌 **String** 이용)     
> **TransactionData = 사용안함. 작업파일 리스트를 이용, 각 파일명을 TransactionItem으로 이용**   
* * *
![get_Book_Sales_Status_Report_result](https://github.com/pnmGithub/get_Book_Sales_Status_Report/assets/149296871/37375cff-3493-4e80-8e79-a5b3c45aea86)
* * *


#### [Config.xlsx 참고] ####
* strZipFilePath - 도서정보 압축 파일 경로
* strExtractedFolder - 도서정보 압축 해제 폴더
* strWorkFolderPath - 작업폴더 (도서별 정보 저장폴더)
* strWorkFileName - 집행품의서 파일명
* strWorkFileTemplate - 도서판매현황 Template 파일
* strResultFolderPath - 결과폴더
* strExcelVBAFile - 엑셀에 반영할 VBA 파일
* strMailAddressFile - 담당자 연락처 파일 (메일주소 세팅 필요)
* strOutlookMailAccount - 아웃룩메일계정
* strMailSubject - 메일 제목
* strMyName - 본인 이름
* strMailBody - 메일 내용


#### [추가된 파일(Invoke) 정보] ####
* Initalization > Bizwork\CreateFolderFile.xaml   
   - 폴더, 파일 처리. 폴더 없으면 생성, 파일 있으면 삭제
* Initalization > Bizwork\GetDictionaryFromExcelReadCell.xaml   
   - 엑셀 데이터를 Read Cell로 받아 Dictionary에 저장하여 반환
* Initalization > Bizwork\GetExtractedFiles.xaml   
   - 압축파일 해제
* Initalization > Bizwork\SetBookSaleReportExcel.xaml   
   - 도서정보 데이터를 각 중분류별로 작업폴더 엑셀파일에 쓰기
* Process > Bizwork\GetWorkExcelFileWriteSheet.xaml   
   - 작업폴더 엑셀파일의 "월별판매현황" 시트에 쓸 데이터테이블 구하기
* End Process > Bizwork\SendMail.xaml   
   - 결과폴더 파일별로 담당자에게 메일발송

     
#### [How It Works] ####

1. **INITIALIZE PROCESS**
   + * 폴더 초기화 - 작업, 결과 폴더 삭제 및 생성
   + 압축파일 해제 ("Data\01_요청\도서판매현황_요청.zip")
   + 도서정보 엑셀파일 가져오기
   + 도서판매현황 파일 생성 - 중분류 별 데이터 쓰기
   + 작업폴더 엑셀파일 가져오기 - arr_TransactionData

2. **GET TRANSACTION DATA**
   + 처리할 TransactionItem 체크

4. **PROCESS TRANSACTION**
   + 작업폴더 엑셀파일 별로
   + * 데이터 없는 중분류시트 삭제
     * 중분류 별 데이터 "판매현황" 시트에 합쳐서 쓰기 - VBA로 시트복사, 시트삭제 사용
     * 판매현황 시트 데이터를 집계해서 "월별판매현황" 시트에 쓰기
     * 결과파일 생성하여 판매현황, 월별판매현황 데이터 쓰기 - VBA로 시트복사, 시트삭제 사용

4. **END PROCESS**
   + 결과 폴더의 파일 가져오기
   + 파일별 담당자에게 메일 발송 (연락처 파일에서 메일주소 가져옴. Lookup Data Table 사용)

* * *
![get_Book_Sales_Status_Report_guide](https://github.com/pnmGithub/get_Book_Sales_Status_Report/assets/149296871/841a03b2-a46a-4616-a4ca-a7e55c2003c3)
* * *

### REFrameWork Template ###
**Robotic Enterprise Framework**
