FROM nginx:alpine
COPY .nginx.conf /etc/nginx/conf.d/default.conf
COPY . /usr/share/nginx/html
RUN rm -rf /usr/share/nginx/html/.git \
           /usr/share/nginx/html/_apps-script \
           /usr/share/nginx/html/.fursys-deploy-hub
EXPOSE 80
HEALTHCHECK --interval=30s --timeout=3s CMD wget -qO- http://127.0.0.1/ || exit 1
CMD ["nginx", "-g", "daemon off;"]
