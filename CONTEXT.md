# Quiz Generator & Evaluator

Generates a randomized MCQ question paper per student from a shared question bank,
then scores the students' Google-Form submissions against the bank's answer key.

## Language

### Papers

**Question Bank**:
The full pool of MCQs a quiz is drawn from. Each entry has a stable **Question Number**.
_Avoid_: question pool, master list

**Question Number**:
A question's identity within the Question Bank (1..N). Stable across every sheet in the
system — it is what an answer column refers to. Printed on a paper as `QCd`, in the
zero-padded `Q- 07` form students copy into the Google Form.
_Avoid_: Q.No, Sr (that is the paper-local position), question id

**Set**:
One student's randomized selection of questions from the Question Bank. Labelled
`S-01`, `S-02`, … — zero-padded, and the same string in the question papers, the Google
Form's set dropdown, and every output sheet.
_Avoid_: Set_1, paper, variant, quiz

**Same Questions mode**:
A way of building Sets where every student gets every question in the Question Bank. Sets
differ only in the order of the questions and the order of each question's A-D options.
_Avoid_: full-bank mode, shuffle-only mode

**Position**:
A question's place within one Set (1st, 2nd, …) together with how its options are
printed there. The same Question Number sits at different Positions in different Sets.
_Avoid_: slot (that is an option's place, A-D), Sr, index

**Set Map**:
The record of every Set's Positions in Same Questions mode: which Question Number is at
each Position, the bank letter printed in each of A-D, and the letter that is Correct
there. Written to the `Set_Map` sheet so answers can be scored per Set.
_Avoid_: key sheet, mapping, layout

**Answer Key**:
The correct option (A/B/C/D) for every question in the Question Bank.
_Avoid_: solution, marking scheme

### Responses

**Response Sheet**:
The raw Google Forms export of student submissions — one row per submission, one answer
column per Question Number, sparse (a student fills only the questions in their Set).
_Avoid_: Google download, submissions, answers file

**Submission**:
One row of the Response Sheet. A student may have several; only the latest one per
Roll Number is scored.
_Avoid_: response, entry, attempt

**Roll Number**:
A student's institutional identity, and the key submissions are deduplicated on.
_Avoid_: student id, student number, roll no

**Attempted**:
Every question a student put an answer against, including Extra Answers. Surfaced to the
faculty as `Count`.
_Avoid_: answered, filled

**Correct**:
The number of a student's answers to questions **in their own Set** that match the Answer
Key. Surfaced to the faculty as `AnsC`. A blank is never Correct, and neither is an Extra
Answer however well it matches the key.
_Avoid_: score, marks obtained, AnsC (that is the column label, not the concept)

**Extra Answer**:
An answer given to a question that was not in the student's Set. Counts towards Attempted,
never towards Correct, and shown on a light yellow background so the distinction is visible.
_Avoid_: invalid answer, out-of-set answer
