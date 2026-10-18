//! Office for the web reads functions.json and shortcuts.json from its own
//! origin with an XHR, so exactly those two answers (a GET or a HEAD) carry
//! `Access-Control-Allow-Origin: *`. Nothing else on this host may: not the
//! panes or their bundles, never `/api/*` (sealed blobs behind a bearer key).
//! Each test builds its own static tree and runs the real router.

mod common;

use std::sync::Arc;

use axum::{
    body::Body,
    http::{header, Method, Request, StatusCode},
    response::Response,
    Router,
};
use common::{static_dir, AUTH, ID};
use plsfix_server::{app, relay::AppState, store::Store};
use tower::ServiceExt;

/// Where a read comes from: Office for the web, in any region.
const OFFICE_ORIGIN: &str = "https://usc-excel.officeapps.live.com";
const METADATA_FILES: [&str; 2] = ["/functions.json", "/shortcuts.json"];

fn app_over(name: &str) -> Router {
    let dir = static_dir(name);
    std::fs::write(dir.join("functions.json"), r#"{"functions":[]}"#).unwrap();
    std::fs::write(dir.join("shortcuts.json"), r#"{"actions":[]}"#).unwrap();
    // The same names one level down: only the two exact paths may match.
    std::fs::write(dir.join("assets").join("functions.json"), "{}").unwrap();
    std::fs::write(dir.join("assets").join("taskpane-abc.js"), "// bundle").unwrap();
    app(dir, Arc::new(AppState::new(Store::in_memory().unwrap())))
}

/// One cross-origin request, the way the browser sends it: with an `Origin`.
async fn ask(app: &Router, method: Method, path: &str) -> Response {
    let request = Request::builder()
        .method(method)
        .uri(path)
        .header(header::ORIGIN, OFFICE_ORIGIN)
        .body(Body::empty())
        .unwrap();
    app.clone().oneshot(request).await.unwrap()
}

fn allow_origin(response: &Response) -> Option<String> {
    response
        .headers()
        .get(header::ACCESS_CONTROL_ALLOW_ORIGIN)
        .map(|value| value.to_str().unwrap().to_string())
}

#[tokio::test]
async fn both_metadata_files_answer_a_cross_origin_read_with_a_star() {
    let app = app_over("cors-both");
    for path in METADATA_FILES {
        for method in [Method::GET, Method::HEAD] {
            let response = ask(&app, method.clone(), path).await;
            assert_eq!(response.status(), StatusCode::OK, "{method} {path}");
            assert_eq!(
                allow_origin(&response).as_deref(),
                Some("*"),
                "{method} {path}"
            );
        }
    }
}

/// The browser may add a query to the URL and revalidates what it cached: the
/// 304 that answers it is still a cross-origin answer.
#[tokio::test]
async fn a_query_string_and_a_revalidation_keep_it() {
    let app = app_over("cors-revalidate");
    for path in METADATA_FILES {
        let with_query = ask(&app, Method::GET, &format!("{path}?v=1")).await;
        assert_eq!(
            allow_origin(&with_query).as_deref(),
            Some("*"),
            "{path}?v=1"
        );

        let first = ask(&app, Method::GET, path).await;
        let stamp = first.headers().get(header::LAST_MODIFIED).unwrap().clone();
        let request = Request::get(path)
            .header(header::ORIGIN, OFFICE_ORIGIN)
            .header(header::IF_MODIFIED_SINCE, stamp)
            .body(Body::empty())
            .unwrap();
        let again = app.clone().oneshot(request).await.unwrap();
        assert_eq!(again.status(), StatusCode::NOT_MODIFIED, "{path}");
        assert_eq!(allow_origin(&again).as_deref(), Some("*"), "{path} 304");
    }
}

#[tokio::test]
async fn the_files_keep_the_no_cache_rule_beside_it() {
    let app = app_over("cors-no-cache");
    for path in METADATA_FILES {
        let response = ask(&app, Method::GET, path).await;
        assert_eq!(
            response.headers().get(header::CACHE_CONTROL).unwrap(),
            "no-cache",
            "{path}"
        );
    }
}

#[tokio::test]
async fn no_other_answer_carries_it() {
    let app = app_over("cors-nothing-else");
    for path in [
        "/",
        "/taskpane.html",
        "/pptpane.html",
        "/assets/taskpane-abc.js",
        "/healthz",
        "/version",
        "/manifest.xml",
        "/nope.json",
    ] {
        let response = ask(&app, Method::GET, path).await;
        assert_eq!(allow_origin(&response), None, "GET {path}");
    }
}

#[tokio::test]
async fn only_the_two_exact_paths_match() {
    let app = app_over("cors-exact");
    // `/assets/functions.json` exists and answers 200: the name alone is not enough.
    let nested = ask(&app, Method::GET, "/assets/functions.json").await;
    assert_eq!(nested.status(), StatusCode::OK);
    assert_eq!(allow_origin(&nested), None, "/assets/functions.json");

    for path in [
        "/api/functions.json",
        "/api/shortcuts.json",
        "/nope/shortcuts.json",
        "/functions.json.map",
        "/functions.jsonx",
        "/Shortcuts.json",
    ] {
        let response = ask(&app, Method::GET, path).await;
        assert_eq!(allow_origin(&response), None, "GET {path}");
    }
}

/// The relay's blobs are sealed and its routes take a bearer key: a read from
/// another origin must never be handed a header that says it may.
#[tokio::test]
async fn the_relay_never_carries_it_even_for_a_cross_origin_read() {
    let app = app_over("cors-relay");
    let link = format!("/api/links/{ID}");

    let push = Request::put(&link)
        .header(header::ORIGIN, OFFICE_ORIGIN)
        .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
        .body(Body::from("sealed"))
        .unwrap();
    let pushed = app.clone().oneshot(push).await.unwrap();
    assert_eq!(pushed.status(), StatusCode::OK);
    assert_eq!(allow_origin(&pushed), None, "PUT {link}");

    let read = Request::get(&link)
        .header(header::ORIGIN, OFFICE_ORIGIN)
        .header(header::AUTHORIZATION, format!("Bearer {AUTH}"))
        .body(Body::empty())
        .unwrap();
    let read = app.clone().oneshot(read).await.unwrap();
    assert_eq!(read.status(), StatusCode::OK);
    assert_eq!(allow_origin(&read), None, "GET {link}");
    assert_eq!(
        read.headers().get(header::CACHE_CONTROL).unwrap(),
        "no-store, private"
    );

    let anonymous = ask(&app, Method::GET, &link).await;
    assert_ne!(anonymous.status(), StatusCode::OK);
    assert_eq!(allow_origin(&anonymous), None, "anonymous GET {link}");
}

/// Only a read is answered for another origin: a write, or a preflight asking
/// leave to send one, gets no header (and no 200: the static host has no such route).
#[tokio::test]
async fn only_a_read_gets_it() {
    let app = app_over("cors-reads-only");
    for path in METADATA_FILES {
        for method in [Method::POST, Method::PUT, Method::DELETE, Method::OPTIONS] {
            let response = ask(&app, method.clone(), path).await;
            assert_ne!(response.status(), StatusCode::OK, "{method} {path}");
            assert_eq!(allow_origin(&response), None, "{method} {path}");
        }
    }
}
