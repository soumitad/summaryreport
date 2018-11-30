# summaryreport

1. Download the project into your workspace. You could use any IDE like Eclipse or IntelliJ
2. Ensure that Maven is installed in your system
3. To create an executable JAR, please run the following command: mvn clean install
4. The above command would create the summaryreport-1.0-SNAPSHOT-distribution.zip, summaryreport-1.0-SNAPSHOT.jar and target classes under target folder. If target folder doesnt get created, please create a target folder under summaryreport

To run the project:

1. Take the summaryreport-1.0-SNAPSHOT-distribution.zip from target folder and place it under C:/Data/Reports
2. The zip contains the executable jar and the dependent jars present under lib folder (which also would be generated)
3. Place the Excel file (Format: Final_BP-ACTIVE-SITE-LIST.xlsx) in the same folder location (C:/Data/Reports)
4. Double click the .bat file which should run the jar file and produce an Output-report.xlsx in the same folder location.