FROM ubuntu:22.04

RUN apt-get -y update && \
    apt-get -y dist-upgrade

RUN apt-get install -y git python3 python3-pip

RUN git clone --depth=1 https://github.com/AVBelyy/AnkiExporter.git
WORKDIR AnkiExporter

RUN pip3 install -r requirements.txt

CMD coffee server.coffee
