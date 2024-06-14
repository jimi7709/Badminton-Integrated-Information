
프로그램 실행 관련해서 첨부된 파일:
 apache-log4j-2.18.0-bin.zip, poi-bin-5.2.3-20220909.zip, 외부 라이브러리 추가해야하는 것(jar)

-  실행환경

해당 프로젝트는 액셀 외부 라이브러리를 사용하기 때문에 다음과 같이 라이브러리를
추가하는 것이 필요하다.

ExcelHome.java 파일이 Main()가 있는 클래스이다.


액셀 외부 라이브러리 추가하는 법(압축파일 해제 기준)

1.해당하는 프로젝트 폴더 우클릭 - 
2. Build Path - 
3. Configure Build Path.. - 
4. Libraries - classpath 클릭
5. 우측 Button 중에서 Add External JARs 클릭
6. lib 안에 있는 모든 JAR, ooxml-lib에 있는 모든 JAR, 
poi-bin-5.2.3 위치에 있는 8개의 JAR 그리고 선택
(해당 jar 파일들을 선택하는 것은 외부 라이브러리 추가해야하는 것(jar)폴더에 있는 
모든 jar을 추가하는 것과 같다.)

7. Appy and Close

+ ( log4j-core-2.18.-.jar 파일도 라이브러리를 불러와서 적용한다.)

 !!같이 첨부된 압축 파일과 자바가 호환이 되지 않는 경우!!

위 압축 파일을 풀어서 위와 같이 다시 라이브러리가 추가해서 실행하거나,

https://poi.apache.org/ 에 접속하여서

https://archive.apache.org/dist/poi/release/src/ 또는

https://archive.apache.org/dist/poi/release/bin/ 중에서 Java 버전을 지원하는 압축 파일을 찾는다.

- 실행 조건

프로젝트 폴더의 src 폴더 안에 Bag.xlsx, Racket,xlsx, Rules,xlsx, Shoes,xlsx, ShuttleCock.xlsx 파일에 꼭 있어야만
해당 프로그램이 실행된다. 액셀 파일들의 위치는 상대경로로 나타내면, ./src 이다.


위 조건을 모두 만족하면 프로그램을 실행할 수 있다.!