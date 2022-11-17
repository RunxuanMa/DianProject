后端Excel处理的功能，该功能在很多实际场景中都能⽤到，
   ⽐如解析前端上传的课 程表、或者是解析前端上传的班级分数表格等

⼀个基本不会编程的⼈如何来完成你的代码的部署运⾏⼯作?


在target目录中找到jar包(名字为maven_springboot-0.0.1-SNAPSHOT.jar)

 配置数据库连接(见下)

进入jar所在的文件夹，使用java -jar命令运行jar，项目就能启动(自行配置环境)

进入如下页面,可进行文档的添加或下载
localhost:8080/myexcel  




添加excel可能遇见的错误: <br>
文件重名，这时你的文件不会上传，请检查文件名字 <br>
某数据过长，请修改异常数据<br>
某数据为空，请修改异常数据  <br>


下载excel:
该页面已列出所有添加在库中的excel，输入想要的excel的完整名称提交即可下载  
  
  






开发记录<br><br>

         
 架构：<br>
 本项目采用传统的springboot框架,control层、service层、dao层经典架构<br>
 
   数据库配置：<br>
     
     在src/main/resources/application.yml 配置数据库连接
     url处设置你的数据库名称
     root处设置用户名
     password处设置密码
     请确保填写正确，否则无法连接上您的数据库！
   
          
   control层：<br>
   
    1.ViewControl
     这里类似项目的控制台，用户在前端可以选择自己是添加Excel还是下载Excel
    2.ExcelControl
     这里调用service层的接口
    3.ExceptionController
     (已弃用)
         
   service层<br>
   
    ExcelService：
    API：
    1.void ReadDataFromExcelAndWriteInToDB(String path, String originalFilename)
     读取control层传过来的Excel文件，调用dao层接口，写入到DB<br><br>
    2.Object[][] readExcel(String path) 
     配合ReadDataFromExcelAndWriteInToDB使用
    3. ArrayList<String> getTableMap()
     被control层调用，在下载文档页面中显示文档表 
    4.HSSFWorkbook createExcel(String excelName) 
     接受control层get传来的文件名，调用dao层接口，返回用户想要的文件
   
   dao层<br>
   
   
    ExcelDAO  
    API：  
    1.ArrayList<String> getTableMap() 返回map表中的所有filename
    
    2.ReadDataFromExcelAndWriteInToDB(Object[][] excelData, String originalFilename)
    将Excel的数据写入数据库中，并在map表中加入 originalFilename与对应的tableName
    
    3.String wrongDataProcess(Object[][] excelData )
     按照需求检查用户传来的数据，检查不过则抛出异常，且数据不会写入数据库中
     (优化：可采用AOP拦截器的形式)
     
    4.writeDataIntoExcel(HSSFWorkbook workbook, String excelName)
    通过service传来的ExcelName在DB中查询是否有对应的文件，有则将数据写入到workbook中，否则抛出自定义异常
    (优化：可采用AOP拦截器的形式)
   
 
 
         
 
 
 maven配置:
 引入 springboot、MySQL、junit等依赖
 
 <dependencies>
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter</artifactId>
        </dependency>
        <dependency>
            <groupId>org.junit.platform</groupId>
            <artifactId>junit-platform-launcher</artifactId>
            <version>1.4.0</version>
            <scope>test</scope>
        </dependency>

        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-test</artifactId>
            <scope>test</scope>
        </dependency>

        <!--数据库驱动-->
        <dependency>
            <groupId>mysql</groupId>
            <artifactId>mysql-connector-java</artifactId>
            <version>5.1.47</version>
        </dependency>



        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-data-jdbc</artifactId>
        </dependency>

        <dependency>
            <groupId>org.projectlombok</groupId>
            <artifactId>lombok</artifactId>
            <version>RELEASE</version>
            <scope>compile</scope>
        </dependency>

        <dependency>
            <groupId>org.springframework</groupId>
            <artifactId>spring-web</artifactId>
            <version>5.3.12</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.17</version>
        </dependency>
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>3.17</version>
        </dependency>

        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-web</artifactId>
        </dependency>
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-jdbc</artifactId>
        </dependency>
        <dependency>
            <groupId>org.springframework.boot</groupId>
            <artifactId>spring-boot-starter-thymeleaf</artifactId>
        </dependency>
    </dependencies>
    
    
    


















