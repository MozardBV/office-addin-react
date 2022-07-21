FROM node:16-bullseye

LABEL Maintainer="Patrick Godschalk <p.godschalk@ienpm.nl>"
LABEL Description "Integratie tussen het Mozard zaaksysteem en Microsoft Office"

WORKDIR /build

COPY package-lock.json package.json ./

RUN npm ci

COPY . .

FROM alpine:3.16.1

RUN apk add --update-cache nodejs \
  && rm -rf /var/cache/apk/*

RUN addgroup -S node \
  && adduser -S node -G node

USER node

RUN mkdir /home/node/code

WORKDIR /home/node/code

COPY --from=0 --chown=node:node /build .

RUN mkdir -p /home/node/code/node_modules/.cache/webpack-dev-server

RUN test -f /home/node/.office-addin-dev-certs/localhost.key && cat /home/node/code/.office-addin-dev-certs/localhost.key > /home/node/code/node_modules/.cache/webpack-dev-server/server.pem || echo "Using Webpack certificate"
RUN test -f /home/node/.office-addin-dev-certs/localhost.crt && cat /home/node/code/.office-addin-dev-certs/localhost.crt >> /home/node/code/node_modules/.cache/webpack-dev-server/server.pem || echo ""
RUN test -f /home/node/.office-addin-dev-certs/ca.crt && cat /home/node/code/.office-addin-dev-certs/ca.crt >> /home/node/code/node_modules/.cache/webpack-dev-server/server.pem || echo ""

EXPOSE 3000

CMD ["node_modules/.bin/webpack", "serve", "--mode", "development", "--https"]

HEALTHCHECK --timeout=300s CMD curl --silent --insecure --fail https://127.0.0.1:3000/
