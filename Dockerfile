FROM mcr.microsoft.com/mssql/server:2022-latest

USER root
RUN mkdir -p /usr/config

# Copiamos los scripts de inicialización al contenedor
COPY init.sql /usr/config/init.sql
COPY entrypoint.sh /usr/config/entrypoint.sh

# Damos permisos de ejecución al script bash
RUN chmod +x /usr/config/entrypoint.sh

USER mssql

# Definimos el entrypoint para que ejecute nuestro script al arrancar
ENTRYPOINT ["/bin/bash", "/usr/config/entrypoint.sh"]
