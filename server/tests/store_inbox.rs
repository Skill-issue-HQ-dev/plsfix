//! The per-workspace inbox row cap at the store level (`store_inbox.rs`).
//! Excel's Push re-announces every link it pushes and each re-post renews the
//! row's 7-day TTL, so rows for links a deck already holds can sit at the cap
//! for good. A full workspace therefore makes room by dropping the POSTING
//! key's own oldest live row; it never drops another key's, and answers Full
//! only when the key has nothing of its own to give up. Everything runs on an
//! in-memory sqlite.

use plsfix_server::store::{Store, INBOX_TTL};
use plsfix_server::store_inbox::{InboxRow, Posted};
use plsfix_server::store_room::Caps;

fn hash(n: u8) -> [u8; 32] {
    [n; 32]
}

/// A store whose workspaces hold at most `rows` live inbox rows.
fn capped(rows: i64) -> Store {
    let mut store = Store::in_memory().unwrap();
    store.set_caps(Caps {
        max_bytes: i64::MAX,
        inbox_rows: rows,
    });
    store
}

fn post(store: &Store, key: u8, id: &str, at: i64) -> Posted {
    store.post_inbox("WS", &hash(key), id, b"x", at).unwrap()
}

fn ids(rows: &[InboxRow]) -> Vec<&str> {
    rows.iter().map(|row| row.id.as_str()).collect()
}

#[test]
fn a_full_workspace_drops_the_posting_keys_oldest_row_for_a_new_one() {
    let store = capped(3);
    for (id, at) in [("a", 10), ("b", 20), ("c", 30)] {
        assert!(matches!(post(&store, 1, id, at), Posted::Stored));
    }

    let posted = post(&store, 1, "d", 40);

    assert!(matches!(posted, Posted::Stored), "{posted:?}");
    let rows = store.list_inbox("WS", &hash(1), 40).unwrap();
    assert_eq!(ids(&rows), ["d", "c", "b"], "the oldest, `a`, is gone");
    assert_eq!(store.inbox_count("WS", 40).unwrap(), 3);
}

#[test]
fn the_oldest_row_is_the_one_posted_longest_ago_not_the_first_inserted() {
    let store = capped(3);
    for (id, at) in [("a", 10), ("b", 20), ("c", 30)] {
        post(&store, 1, id, at);
    }
    // Excel pushes `a` again: the replace renews its time, so `b` is now the
    // row posted longest ago.
    assert!(matches!(post(&store, 1, "a", 40), Posted::Stored));

    assert!(matches!(post(&store, 1, "d", 50), Posted::Stored));

    let rows = store.list_inbox("WS", &hash(1), 50).unwrap();
    assert_eq!(ids(&rows), ["d", "a", "c"]);
}

#[test]
fn rows_of_one_second_give_way_in_insertion_order() {
    let store = capped(2);
    post(&store, 1, "first", 10);
    post(&store, 1, "second", 10);

    assert!(matches!(post(&store, 1, "third", 10), Posted::Stored));

    let rows = store.list_inbox("WS", &hash(1), 10).unwrap();
    assert_eq!(ids(&rows), ["third", "second"]);
}

#[test]
fn a_replace_at_the_cap_drops_nothing() {
    let store = capped(3);
    for (id, at) in [("a", 10), ("b", 20), ("c", 30)] {
        post(&store, 1, id, at);
    }

    assert!(matches!(post(&store, 1, "b", 40), Posted::Stored));

    let rows = store.list_inbox("WS", &hash(1), 40).unwrap();
    assert_eq!(ids(&rows), ["b", "c", "a"], "all three rows still here");
}

#[test]
fn a_workspace_another_key_filled_stays_full_for_a_key_with_no_rows_of_its_own() {
    let store = capped(3);
    for id in ["x", "y", "z"] {
        post(&store, 2, id, 10);
    }

    let posted = post(&store, 1, "mine", 20);

    assert!(matches!(posted, Posted::Full), "{posted:?}");
    // The other key's rows are untouched, and this key still sees nothing.
    let theirs = store.list_inbox("WS", &hash(2), 20).unwrap();
    assert_eq!(ids(&theirs), ["z", "y", "x"]);
    assert!(store.list_inbox("WS", &hash(1), 20).unwrap().is_empty());
}

#[test]
fn a_key_gives_up_only_its_own_row_when_other_keys_hold_the_rest() {
    let store = capped(3);
    post(&store, 2, "x", 10);
    post(&store, 2, "y", 11);
    post(&store, 1, "mine-1", 12);

    let posted = post(&store, 1, "mine-2", 13);

    assert!(matches!(posted, Posted::Stored), "{posted:?}");
    let mine = store.list_inbox("WS", &hash(1), 13).unwrap();
    assert_eq!(ids(&mine), ["mine-2"], "its own oldest row made the room");
    let theirs = store.list_inbox("WS", &hash(2), 13).unwrap();
    assert_eq!(ids(&theirs), ["y", "x"], "the other key lost nothing");
}

#[test]
fn another_workspace_is_unaffected_by_a_full_one() {
    let store = capped(1);
    post(&store, 1, "a", 10);

    let elsewhere = store.post_inbox("OTHER", &hash(1), "a", b"x", 11).unwrap();

    assert!(matches!(elsewhere, Posted::Stored), "{elsewhere:?}");
    assert_eq!(store.inbox_count("WS", 11).unwrap(), 1);
    assert_eq!(store.inbox_count("OTHER", 11).unwrap(), 1);
}

#[test]
fn an_expired_row_makes_room_before_any_live_one_is_dropped() {
    let store = capped(3);
    for (id, at) in [("a", 0), ("b", 1), ("c", 2)] {
        post(&store, 1, id, at);
    }

    // `a` reaches its 7 days exactly now and is swept (a row dies at
    // `expires_at <= now`); `b` and `c` are live, so two rows leave room for one.
    let posted = post(&store, 1, "d", INBOX_TTL);

    assert!(matches!(posted, Posted::Stored), "{posted:?}");
    let rows = store.list_inbox("WS", &hash(1), INBOX_TTL).unwrap();
    assert_eq!(ids(&rows), ["d", "c", "b"], "nothing live was dropped");
}

#[test]
fn the_byte_ceiling_refuses_before_anything_is_dropped() {
    let mut store = Store::in_memory().unwrap();
    store.set_caps(Caps {
        max_bytes: 3,
        inbox_rows: 3,
    });
    for (id, at) in [("a", 10), ("b", 20), ("c", 30)] {
        post(&store, 1, id, at);
    }

    // Three one-byte rows fill the ceiling; a fourth byte does not fit, and
    // dropping a row to make a count of rows fit would be the wrong answer.
    let posted = post(&store, 1, "d", 40);

    assert!(matches!(posted, Posted::Full), "{posted:?}");
    let rows = store.list_inbox("WS", &hash(1), 40).unwrap();
    assert_eq!(ids(&rows), ["c", "b", "a"]);
}
