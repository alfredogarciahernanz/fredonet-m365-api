# Fredonet M365 API

API FastAPI para integración con Microsoft 365 y Microsoft Graph.

## Características

- Microsoft Graph API
- Calendar Integration
- Mail Integration
- API Key Authentication
- FastAPI
- Docker Compose

## Endpoints

### Health

GET /health

### Version

GET /version

## Despliegue

docker compose up -d

## Variables de entorno

Ver .env.example

## Seguridad

- API Key Authentication
- Variables sensibles mediante .env
- Secretos excluidos de Git
- Timeouts explícitos
- Validación defensiva de respuestas Graph
