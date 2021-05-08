FROM python:3.6-slim-stretch
#E: Unable to locate package libjasper-dev

ARG basemap_version=master

WORKDIR /TCC

RUN apt-get update \
            && apt-get -y upgrade \
            && apt-get -y install --no-install-recommends \
            make \
            curl \
            wget \
            nano \
            && apt-get -y install --no-install-recommends \
            vim \
            python3-pip \
            python3-matplotlib \
            python3-setuptools \
            python3-numpy \
            python3-dev \
            build-essential \
            zlib1g \
            zip \
	    bc \
	    git \
	    iproute2 \
	    bzip2 \
	    python3-pip \
	    python3-setuptools \
	    ssh \
	    build-essential \
	    && pip3 install --upgrade pip \
	    && pip3 install simplejson \
	    && pip3 install ipython \
	    && pip3 install pathlib \
	    && pip3 install scipy \
	    && pip3 install numpy \
        && pip3 install pandas \
	    && pip3 install python-dateutil \
	    && pip3 install plotly \
	    && pip3 install sklearn \
        && pip3 install seaborn \
        && pip3 install pystan==2.19.1.1 \
            && pip3 install prophet

COPY . .        

ENTRYPOINT ["python"]