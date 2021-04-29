FROM openjdk:11-jre-slim
RUN java -version 

RUN groupadd -g 1700 gisadmin && \
  useradd -ms /bin/bash -u 17000 -g 1700 gisadmin 

ADD https://github.com/rtrier/excel2db/raw/master/de.gdiservice.excel2db/lib/de.gdiservice.excel2db.jar /gdiservice-lib/main.jar
ADD archiv.properties /gdiservice-lib/


WORKDIR /exceldata/

ENTRYPOINT [ "java", "-cp", "/gdiservice-lib/:/gdiservice-lib/main.jar", "de.gdiservice.excel2db.Excel2DB" ]

