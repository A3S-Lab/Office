use std::sync::Arc;

use a3s_boot::{
    AxumAdapter, BootApplication, BootError, BootErrorKind, BootRequest, BootResponse,
    ConfigModule, ControllerDefinition, Module, ModuleRef, ProviderDefinition, Result, Validate,
    WebSocketContext, WebSocketExceptionResponse, WebSocketGatewayDefinition, WebSocketMessage,
};
use a3s_office::{
    NATIVE_OFFICE_COLLABORATION_PROTOCOL, NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};
use serde::Serialize;
use serde_json::json;
use time::format_description::well_known::Rfc3339;
use time::OffsetDateTime;

use crate::config::CollaborationConfig;
use crate::protocol::{TicketRequest, AWARENESS_EVENT, DOCUMENT_EVENT, ERROR_EVENT, HELLO_EVENT};
use crate::service::CollaborationService;

#[derive(Debug)]
struct CollaborationModule {
    config: ConfigModule<CollaborationConfig>,
}

impl CollaborationModule {
    fn new(config: CollaborationConfig) -> Self {
        Self {
            config: ConfigModule::from_value("collaboration-config", config),
        }
    }
}

impl Module for CollaborationModule {
    fn name(&self) -> &'static str {
        "office-collaboration"
    }

    fn imports(&self) -> Vec<Arc<dyn Module>> {
        vec![Arc::new(self.config.clone())]
    }

    fn providers(&self) -> Result<Vec<ProviderDefinition>> {
        Ok(vec![
            ProviderDefinition::factory::<CollaborationService, _>(|module_ref| {
                let config = module_ref.get::<CollaborationConfig>()?;
                Ok(CollaborationService::new((*config).clone()))
            }),
        ])
    }

    fn controllers(&self, module_ref: &ModuleRef) -> Result<Vec<ControllerDefinition>> {
        let health = ControllerDefinition::new("/api/collaboration")?.get(
            "/healthz",
            |request: BootRequest| async move {
                api_success(
                    200,
                    request_id(&request),
                    json!({
                        "status": "ok",
                        "protocol": NATIVE_OFFICE_COLLABORATION_PROTOCOL,
                        "version": NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
                    }),
                )
            },
        )?;

        let service = module_ref.get::<CollaborationService>()?;
        let tickets = ControllerDefinition::new("/api/collaboration")?.post(
            "/tickets",
            move |request: BootRequest| {
                let service = Arc::clone(&service);
                async move {
                    let request_id = request_id(&request);
                    let result = request.json::<TicketRequest>().and_then(|body| {
                        service.issue_ticket(request.header("authorization"), body)
                    });
                    match result {
                        Ok(ticket) => api_success(201, request_id, ticket),
                        Err(error) => api_error(request_id, error),
                    }
                }
            },
        )?;
        Ok(vec![health, tickets])
    }

    fn gateways(&self, module_ref: &ModuleRef) -> Result<Vec<WebSocketGatewayDefinition>> {
        let service = module_ref.get::<CollaborationService>()?;
        let connect_service = Arc::clone(&service);
        let disconnect_service = Arc::clone(&service);
        let hello_service = Arc::clone(&service);
        let document_service = Arc::clone(&service);
        let awareness_service = Arc::clone(&service);
        let gateway =
            WebSocketGatewayDefinition::new("/collaboration/{artifactKind}/{artifactId}")?
                .with_connection_hook(move |connection| {
                    let service = Arc::clone(&connect_service);
                    async move { service.authenticate_connection(connection).await }
                })
                .with_disconnect_hook(move |connection| {
                    let service = Arc::clone(&disconnect_service);
                    async move { service.disconnect(connection).await }
                })
                .subscribe_with_connection(HELLO_EVENT, move |connection, message| {
                    let service = Arc::clone(&hello_service);
                    async move { service.hello(connection, message).await }
                })?
                .subscribe_with_connection(DOCUMENT_EVENT, move |connection, message| {
                    let service = Arc::clone(&document_service);
                    async move { service.receive_document(connection, message).await }
                })?
                .subscribe_with_connection(AWARENESS_EVENT, move |connection, message| {
                    let service = Arc::clone(&awareness_service);
                    async move { service.receive_awareness(connection, message).await }
                })?
                .subscribe("collaboration.ping", |_| async {
                    Ok(WebSocketMessage::new(
                        "collaboration.pong",
                        json!({ "protocol": NATIVE_OFFICE_COLLABORATION_PROTOCOL }),
                    ))
                })?
                .with_filter(|context: WebSocketContext, error: BootError| async move {
                    Ok(Some(WebSocketExceptionResponse::message(
                        WebSocketMessage::new(
                            ERROR_EVENT,
                            json!({
                                "code": business_error_code(error.kind()),
                                "message": error.http_response_message(),
                                "event": context.event,
                            }),
                        ),
                    )))
                });
        Ok(vec![gateway])
    }
}

pub fn build_application(config: CollaborationConfig) -> Result<BootApplication> {
    config.validate()?;
    BootApplication::builder()
        .import(CollaborationModule::new(config))
        .build()
}

pub async fn serve(config: CollaborationConfig) -> Result<()> {
    let bind = config.bind;
    let app = build_application(config)?;
    app.serve_with(&AxumAdapter::new(), bind).await
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ApiSuccess<T> {
    code: u16,
    message: &'static str,
    data: T,
    request_id: String,
    timestamp: String,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct ApiFailure {
    code: u16,
    status_code: &'static str,
    message: String,
    details: serde_json::Value,
    request_id: String,
    timestamp: String,
}

fn api_success<T: Serialize>(status: u16, request_id: String, data: T) -> Result<BootResponse> {
    BootResponse::json_with_status(
        status,
        &ApiSuccess {
            code: status,
            message: "Success",
            data,
            request_id,
            timestamp: timestamp(),
        },
    )
}

fn api_error(request_id: String, error: BootError) -> Result<BootResponse> {
    let status = error.http_status_code();
    BootResponse::json_with_status(
        status,
        &ApiFailure {
            code: status,
            status_code: business_error_code(error.kind()),
            message: error.http_response_message(),
            details: json!({}),
            request_id,
            timestamp: timestamp(),
        },
    )
}

fn business_error_code(kind: BootErrorKind) -> &'static str {
    match kind {
        BootErrorKind::BadRequest => "BAD_REQUEST",
        BootErrorKind::Unauthorized => "UNAUTHORIZED",
        BootErrorKind::Forbidden => "FORBIDDEN",
        BootErrorKind::NotFound => "NOT_FOUND",
        BootErrorKind::Conflict => "CONFLICT",
        BootErrorKind::PayloadTooLarge => "PAYLOAD_TOO_LARGE",
        BootErrorKind::ServiceUnavailable => "SERVICE_UNAVAILABLE",
        _ => "INTERNAL_SERVER_ERROR",
    }
}

fn request_id(request: &BootRequest) -> String {
    if let Some(value) = request.header("x-request-id") {
        let value = value.trim();
        if !value.is_empty() && value.len() <= 128 && value.is_ascii() {
            return value.to_string();
        }
    }
    let mut bytes = [0_u8; 16];
    if getrandom::fill(&mut bytes).is_err() {
        return format!("request-{}", std::process::id());
    }
    format!(
        "{:02x}{:02x}{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        bytes[0],
        bytes[1],
        bytes[2],
        bytes[3],
        bytes[4],
        bytes[5],
        bytes[6],
        bytes[7],
        bytes[8],
        bytes[9],
        bytes[10],
        bytes[11],
        bytes[12],
        bytes[13],
        bytes[14],
        bytes[15],
    )
}

fn timestamp() -> String {
    OffsetDateTime::now_utc()
        .format(&Rfc3339)
        .unwrap_or_else(|_| "1970-01-01T00:00:00Z".to_string())
}
