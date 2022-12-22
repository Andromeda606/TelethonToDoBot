FROM ubuntu:20.04
RUN apt-get update
RUN apt-get install python3-pip -y
RUN apt install lsof
RUN apt clean
RUN apt-get install language-pack-tr
ADD ./ ./
RUN pip install -r req.txt