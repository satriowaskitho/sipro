FROM rocker/r-ver:4.3.0

# Install Shiny Server
RUN apt-get update && apt-get install -y \
    gdebi-core \
    wget

RUN wget https://download3.rstudio.org/ubuntu-18.04/x86_64/shiny-server-1.5.20.1002-amd64.deb
RUN gdebi -n shiny-server-1.5.20.1002-amd64.deb

# Copy renv files
COPY renv.lock renv.lock
COPY .Rprofile .Rprofile
COPY renv/activate.R renv/activate.R

# Restore R packages from renv
RUN R -e "renv::restore()"

# Copy app
COPY ./ /srv/shiny-server/app

# Create directory for credentials
RUN mkdir -p /srv/shiny-server/app/credentials

EXPOSE 3838

# Decode base64 secret and run Shiny
CMD echo $GS4_SA_JSON_BASE64 | base64 -d > /srv/shiny-server/app/credentials/gs4-sa.json && \
    /usr/bin/shiny-server
