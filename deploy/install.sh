#!/bin/bash
# install.sh — instala la app en un VPS Ubuntu 22.04 limpio
# Uso: sudo bash install.sh TU_DOMINIO
# Requisito: el DNS del dominio ya debe apuntar a este servidor

set -e
DOMAIN="${1:?Uso: sudo bash install.sh TU_DOMINIO}"
APP_DIR="/opt/sri_conciliacion"
APP_USER="sri"

echo "=== [1/7] Actualizando sistema ==="
apt-get update -q && apt-get upgrade -y -q

echo "=== [2/7] Instalando dependencias del sistema ==="
apt-get install -y -q python3 python3-pip python3-venv nginx certbot python3-certbot-nginx git

echo "=== [3/7] Creando usuario de sistema '$APP_USER' ==="
id -u "$APP_USER" &>/dev/null || useradd --system --create-home --shell /bin/bash "$APP_USER"

echo "=== [4/7] Clonando / copiando código ==="
mkdir -p "$APP_DIR"
# Si hay repositorio git:
# git clone <repo_url> "$APP_DIR"
# Si se copia manualmente (scp), este paso ya está hecho.
chown -R "$APP_USER":"$APP_USER" "$APP_DIR"

echo "=== [5/7] Instalando dependencias Python ==="
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/venv/bin/pip" install --quiet -r "$APP_DIR/requirements.txt"

echo "=== [6/7] Configurando nginx y TLS ==="
mkdir -p /var/log/sri_conciliacion
chown "$APP_USER":"$APP_USER" /var/log/sri_conciliacion
sed "s/TU_DOMINIO/$DOMAIN/g" "$APP_DIR/deploy/nginx.conf" > /etc/nginx/sites-available/sri_conciliacion
ln -sf /etc/nginx/sites-available/sri_conciliacion /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx
certbot --nginx -d "$DOMAIN" --non-interactive --agree-tos -m "admin@$DOMAIN"

echo "=== [7/7] Configurando servicio systemd ==="
cp "$APP_DIR/deploy/sri-conciliacion.service" /etc/systemd/system/
systemctl daemon-reload
systemctl enable sri-conciliacion
systemctl start sri-conciliacion

echo ""
echo "Instalación completa."
echo "   App corriendo en: https://$DOMAIN"
echo ""
echo "Próximo paso — añadir el primer usuario:"
echo "   sudo -u $APP_USER $APP_DIR/venv/bin/python $APP_DIR/setup_usuario.py add <nombre>"
