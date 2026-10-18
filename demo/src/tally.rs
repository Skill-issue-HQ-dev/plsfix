// What a sheet builder wrote: labels, numbers, formulas and formatted blanks,
// counted at the write site so the saved file can be reconciled against the
// builders' own count instead of eyeballed.

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct Tally {
    pub strings: usize,
    pub numbers: usize,
    pub formulas: usize,
    /// Cells with a format and no value (the stray formatting Clean past the data removes).
    pub blanks: usize,
}

impl Tally {
    pub fn cells(&self) -> usize {
        self.strings + self.numbers + self.formulas + self.blanks
    }
}

impl std::ops::AddAssign for Tally {
    fn add_assign(&mut self, other: Tally) {
        self.strings += other.strings;
        self.numbers += other.numbers;
        self.formulas += other.formulas;
        self.blanks += other.blanks;
    }
}
