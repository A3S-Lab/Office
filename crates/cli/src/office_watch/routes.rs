use std::convert::Infallible;
use std::sync::Arc;
use std::time::Duration;

use axum::body::{Body, Bytes};
use axum::extract::{Request, State};
use axum::http::header::{
    CACHE_CONTROL, CONTENT_SECURITY_POLICY, CONTENT_TYPE, COOKIE, HOST, REFERRER_POLICY,
    SET_COOKIE, X_CONTENT_TYPE_OPTIONS, X_FRAME_OPTIONS,
};
use axum::http::{HeaderMap, HeaderName, HeaderValue, StatusCode};
use axum::middleware::{self, Next};
use axum::response::sse::{Event, KeepAlive, Sse};
use axum::response::{IntoResponse, Response};
use axum::routing::get;
use axum::{Json, Router};
use futures_util::stream::{self, Stream, StreamExt};

use super::page::{WATCH_PAGE, WATCH_SCRIPT};
use super::{NativeOfficeWatchStatus, WatchState};

const COOKIE_NAME: &str = "a3s-office-watch";

pub(super) fn router(state: Arc<WatchState>) -> Router {
    Router::new()
        .route("/", get(index))
        .route("/watch.js", get(script))
        .route("/preview", get(preview))
        .route("/status", get(status))
        .route("/events", get(events))
        .fallback(not_found)
        .layer(middleware::from_fn_with_state(
            Arc::clone(&state),
            authorize,
        ))
        .with_state(state)
}

async fn authorize(State(state): State<Arc<WatchState>>, request: Request, next: Next) -> Response {
    let host = request
        .headers()
        .get(HOST)
        .and_then(|value| value.to_str().ok());
    if !host.is_some_and(|host| host_matches(host, state.address))
        || !request_has_token(&request, &state.token)
    {
        return not_found().await;
    }
    next.run(request).await
}

fn host_matches(host: &str, address: std::net::SocketAddr) -> bool {
    host == address.to_string() || (address.port() == 80 && host == address.ip().to_string())
}

fn request_has_token(request: &Request, expected: &str) -> bool {
    let query_matches = request
        .uri()
        .query()
        .into_iter()
        .flat_map(|query| url::form_urlencoded::parse(query.as_bytes()))
        .filter(|(name, _)| name == "token")
        .any(|(_, value)| constant_time_eq(value.as_bytes(), expected.as_bytes()));
    if query_matches {
        return true;
    }
    request
        .headers()
        .get_all(COOKIE)
        .iter()
        .filter_map(|value| value.to_str().ok())
        .flat_map(|value| value.split(';'))
        .filter_map(|value| value.trim().split_once('='))
        .any(|(name, value)| {
            name == COOKIE_NAME && constant_time_eq(value.as_bytes(), expected.as_bytes())
        })
}

fn constant_time_eq(left: &[u8], right: &[u8]) -> bool {
    let mut different = left.len() ^ right.len();
    for index in 0..left.len().max(right.len()) {
        let left = left.get(index).copied().unwrap_or_default();
        let right = right.get(index).copied().unwrap_or_default();
        different |= usize::from(left ^ right);
    }
    different == 0
}

async fn index(State(state): State<Arc<WatchState>>) -> Response {
    let mut headers = document_headers("text/html; charset=utf-8");
    headers.insert(
        CONTENT_SECURITY_POLICY,
        HeaderValue::from_static("default-src 'none'; frame-src 'self'; script-src 'self'; style-src 'unsafe-inline'; connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'; frame-ancestors 'none'"),
    );
    let cookie = format!(
        "{COOKIE_NAME}={}; HttpOnly; SameSite=Strict; Path=/",
        state.token
    );
    if let Ok(cookie) = HeaderValue::from_str(&cookie) {
        headers.insert(SET_COOKIE, cookie);
    }
    (headers, WATCH_PAGE).into_response()
}

async fn script() -> Response {
    (
        document_headers("text/javascript; charset=utf-8"),
        WATCH_SCRIPT,
    )
        .into_response()
}

async fn preview(State(state): State<Arc<WatchState>>) -> Response {
    let html = state.html().await;
    let mut response = Response::new(Body::from(html));
    *response.headers_mut() = document_headers("text/html; charset=utf-8");
    response
}

async fn status(State(state): State<Arc<WatchState>>) -> Response {
    (api_headers(), Json(state.status().await)).into_response()
}

async fn events(
    State(state): State<Arc<WatchState>>,
) -> (
    HeaderMap,
    Sse<impl Stream<Item = Result<Event, Infallible>>>,
) {
    let receiver = state.events.subscribe();
    let shutdown = state.shutdown.subscribe();
    let initial = state.status().await;
    let first = stream::once(async move { Ok(event(initial)) });
    let updates = stream::unfold(
        (receiver, shutdown),
        |(mut receiver, mut shutdown)| async move {
            loop {
                tokio::select! {
                    received = receiver.recv() => match received {
                        Ok(status) => {
                            return Some((Ok(event(status)), (receiver, shutdown)));
                        }
                        Err(tokio::sync::broadcast::error::RecvError::Lagged(_)) => continue,
                        Err(tokio::sync::broadcast::error::RecvError::Closed) => return None,
                    },
                    _ = shutdown.recv() => return None,
                }
            }
        },
    );
    (
        api_headers(),
        Sse::new(first.chain(updates)).keep_alive(
            KeepAlive::new()
                .interval(Duration::from_secs(15))
                .text("keep-alive"),
        ),
    )
}

fn event(status: NativeOfficeWatchStatus) -> Event {
    let name = if status.healthy {
        "snapshot"
    } else {
        "render-error"
    };
    let id = status.version.to_string();
    let data = serde_json::to_string(&status).unwrap_or_else(|_| {
        "{\"healthy\":false,\"version\":0,\"error\":{\"code\":\"use.office.watch_event_failed\",\"message\":\"Failed to serialize watch state.\"}}".to_string()
    });
    Event::default().event(name).id(id).data(data)
}

fn document_headers(content_type: &'static str) -> HeaderMap {
    let mut headers = api_headers();
    headers.insert(CONTENT_TYPE, HeaderValue::from_static(content_type));
    headers.insert(X_FRAME_OPTIONS, HeaderValue::from_static("SAMEORIGIN"));
    headers
}

fn api_headers() -> HeaderMap {
    let mut headers = HeaderMap::new();
    headers.insert(CACHE_CONTROL, HeaderValue::from_static("no-store"));
    headers.insert(X_CONTENT_TYPE_OPTIONS, HeaderValue::from_static("nosniff"));
    headers.insert(REFERRER_POLICY, HeaderValue::from_static("no-referrer"));
    headers.insert(
        HeaderName::from_static("cross-origin-opener-policy"),
        HeaderValue::from_static("same-origin"),
    );
    headers.insert(
        HeaderName::from_static("cross-origin-resource-policy"),
        HeaderValue::from_static("same-origin"),
    );
    headers.insert(
        HeaderName::from_static("permissions-policy"),
        HeaderValue::from_static("camera=(), microphone=(), geolocation=()"),
    );
    headers
}

async fn not_found() -> Response {
    let mut response = Response::new(Body::from(Bytes::from_static(b"Not Found")));
    *response.status_mut() = StatusCode::NOT_FOUND;
    *response.headers_mut() = document_headers("text/plain; charset=utf-8");
    response
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn token_comparison_checks_value_and_length() {
        assert!(constant_time_eq(b"abcdef", b"abcdef"));
        assert!(!constant_time_eq(b"abcdeg", b"abcdef"));
        assert!(!constant_time_eq(b"abcdef0", b"abcdef"));
        assert!(!constant_time_eq(b"abcde", b"abcdef"));
    }

    #[test]
    fn host_validation_accepts_only_the_bound_loopback_authority() {
        let address = "127.0.0.1:39123".parse().unwrap();
        assert!(host_matches("127.0.0.1:39123", address));
        assert!(!host_matches("localhost:39123", address));
        assert!(!host_matches("example.invalid", address));
        assert!(host_matches("127.0.0.1", "127.0.0.1:80".parse().unwrap()));
    }
}
