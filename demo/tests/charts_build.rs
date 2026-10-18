// The Charts sheet must not break the demo's byte-reproducible build: the video's provenance
// check compares the demo's sha256 with the one its screens were captured on, so two builds of
// the same source have to hash alike. Carries its own SHA-256: the demo has no hashing crate.

use std::time::Duration;

/// Longer than the one-second resolution of a document's timestamp.
const CLOCK_TICK: Duration = Duration::from_millis(1100);

/// The first 32 bits of the fractional parts of the cube roots of the first 64 primes (FIPS 180-4).
const K: [u32; 64] = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];

/// The starting state: the square roots of the first eight primes.
const START: [u32; 8] = [
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
];

/// One 64-byte block folded into the state.
fn compress(state: &mut [u32; 8], block: &[u8]) {
    let mut w = [0u32; 64];
    for (slot, word) in w.iter_mut().zip(block.chunks(4)) {
        *slot = u32::from_be_bytes([word[0], word[1], word[2], word[3]]);
    }
    for i in 16..64 {
        let s0 = w[i - 15].rotate_right(7) ^ w[i - 15].rotate_right(18) ^ (w[i - 15] >> 3);
        let s1 = w[i - 2].rotate_right(17) ^ w[i - 2].rotate_right(19) ^ (w[i - 2] >> 10);
        w[i] = w[i - 16]
            .wrapping_add(s0)
            .wrapping_add(w[i - 7])
            .wrapping_add(s1);
    }
    let [mut a, mut b, mut c, mut d, mut e, mut f, mut g, mut h] = *state;
    for (k, word) in K.iter().zip(w) {
        let sigma1 = e.rotate_right(6) ^ e.rotate_right(11) ^ e.rotate_right(25);
        let t1 = h
            .wrapping_add(sigma1)
            .wrapping_add((e & f) ^ (!e & g))
            .wrapping_add(*k);
        let t1 = t1.wrapping_add(word);
        let sigma0 = a.rotate_right(2) ^ a.rotate_right(13) ^ a.rotate_right(22);
        let t2 = sigma0.wrapping_add((a & b) ^ (a & c) ^ (b & c));
        (h, g, f, e, d, c, b, a) = (g, f, e, d.wrapping_add(t1), c, b, a, t1.wrapping_add(t2));
    }
    for (slot, value) in state.iter_mut().zip([a, b, c, d, e, f, g, h]) {
        *slot = slot.wrapping_add(value);
    }
}

/// The SHA-256 of `data` as lowercase hex.
fn sha256(data: &[u8]) -> String {
    let mut message = data.to_vec();
    message.push(0x80);
    while message.len() % 64 != 56 {
        message.push(0);
    }
    message.extend_from_slice(&(data.len() as u64 * 8).to_be_bytes());
    let mut state = START;
    for block in message.chunks(64) {
        compress(&mut state, block);
    }
    state.iter().map(|word| format!("{word:08x}")).collect()
}

fn built() -> Vec<u8> {
    smt_demo::build()
        .expect("build")
        .0
        .save_to_buffer()
        .expect("save")
}

// The FIPS 180-4 examples: the empty message, "abc", and the 448-bit message that needs a
// second block for its padding.
#[test]
fn charts_build_sha256_matches_the_standard_vectors() {
    let long = b"abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq";
    assert_eq!(
        sha256(b""),
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    );
    assert_eq!(
        sha256(b"abc"),
        "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    );
    assert_eq!(
        sha256(long),
        "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
    );
}

// A clock that leaked into the file would show only once the second changes, so the two builds
// wait out one full tick between them.
#[test]
fn charts_build_is_byte_reproducible() {
    let first = sha256(&built());
    std::thread::sleep(CLOCK_TICK);
    let second = sha256(&built());
    assert_eq!(
        first, second,
        "two builds differ: a clock or a random order leaked into the file"
    );
    assert_eq!(first.len(), 64);
}
