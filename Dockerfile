FROM rocker/shiny:latest

# Install system dependencies required by R packages
RUN apt-get update && apt-get install -y \
    libcurl4-gnutls-dev \
    libssl-dev \
    libxml2-dev \
    libfontconfig1-dev \
    libharfbuzz-dev \
    libfribidi-dev \
    libfreetype6-dev \
    libpng-dev \
    libtiff5-dev \
    libjpeg-dev \
    libxt-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

# Install R packages
RUN R -e "install.packages(c( \
    'bs4Dash', \
    'writexl', \
    'dplyr', \
    'DT', \
    'zip', \
    'rlang', \
    'openxlsx', \
    'shinyjs', \
    'shinyalert', \
    'haven', \
    'googlesheets4' \
    ), repos='https://cloud.r-project.org/', dependencies=TRUE)"

# Create directory for credentials
RUN mkdir -p /srv/shiny-server/credentials

# Copy your Shiny app (app.R and www/ folder)
COPY app.R /srv/shiny-server/app.R
COPY www/ /srv/shiny-server/www/

# Expose Shiny Server port
EXPOSE 3838

# Decode secret and run Shiny Server
CMD echo $GS4_SA_JSON_BASE64 | base64 -d > /srv/shiny-server/credentials/gs4-sa.json && \
    /usr/bin/shiny-server