#!/bin/sh
set -e

TEMPLATE="/usr/share/nginx/html/runtime-env.js.template"
TARGET="/usr/share/nginx/html/runtime-env.js"

if [ -f "$TEMPLATE" ]; then
  envsubst '$VITE_AZURE_CLIENT_ID $VITE_AZURE_TENANT_ID $VITE_REDIRECT_URI' < "$TEMPLATE" > "$TARGET"
fi

if [ "$#" -eq 0 ]; then
  set -- nginx -g "daemon off;"
fi

exec "$@"
