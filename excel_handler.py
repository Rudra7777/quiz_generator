"""
Excel Handler Module

Provides functions for:
- Loading question bank from Excel file
- Generating question papers as multi-sheet Excel
"""

import pandas as pd
from dataclasses import dataclass
from typing import List, Dict, Optional
from pathlib import Path


@dataclass
class FullQuestion:
    """
    Represents a complete question with all details.
    """
    question_no: int
    question_id: str  # Internal ID for allocation (e.g., H1, M5, E3)
    question_text: str
    option_a: str
    option_b: str
    option_c: str
    option_d: str
    answer: str  # A, B, C, or D
    difficulty: str  # 'hard', 'medium', 'easy'


class FullQuestionBank:
    """
    Stores complete questions with all metadata.
    """
    
    def __init__(self, questions: List[FullQuestion]):
        self.questions = questions
        self._by_id: Dict[str, FullQuestion] = {q.question_id: q for q in questions}
        self._by_difficulty: Dict[str, List[FullQuestion]] = {
            'hard': [], 'medium': [], 'easy': []
        }
        for q in questions:
            self._by_difficulty[q.difficulty].append(q)
    
    def get_by_id(self, question_id: str) -> Optional[FullQuestion]:
        """Get a question by its ID."""
        return self._by_id.get(question_id)
    
    def get_by_difficulty(self, difficulty: str) -> List[FullQuestion]:
        """Get all questions of a specific difficulty."""
        return self._by_difficulty.get(difficulty, [])
    
    def get_all(self) -> List[FullQuestion]:
        """Get all questions."""
        return self.questions
    
    def count_by_difficulty(self) -> Dict[str, int]:
        """Get count of questions per difficulty."""
        return {d: len(qs) for d, qs in self._by_difficulty.items()}
    
    def get_question_ids_by_difficulty(self, difficulty: str) -> List[str]:
        """Get list of question IDs for a difficulty level."""
        return [q.question_id for q in self._by_difficulty.get(difficulty, [])]


def normalize_difficulty(value: str) -> str:
    """
    Normalize difficulty value to standard format.
    
    Accepts: H/Hard/high, M/Medium/med, L/Low/Easy/easy
    Returns: 'hard', 'medium', or 'easy'
    """
    v = str(value).strip().upper()
    if v in ('H', 'HARD', 'HIGH'):
        return 'hard'
    elif v in ('M', 'MEDIUM', 'MED'):
        return 'medium'
    elif v in ('L', 'LOW', 'E', 'EASY'):
        return 'easy'
    else:
        raise ValueError(f"Unknown difficulty: {value}. Use H/M/L or Hard/Medium/Easy")


def load_question_bank(filepath: str) -> FullQuestionBank:
    """
    Load question bank from Excel file.
    
    Expected columns:
        - question_no: Unique question number
        - question: Question text
        - option_a, option_b, option_c, option_d: Options
        - answer: Correct answer (A/B/C/D)
        - difficulty: H/M/L or Hard/Medium/Easy
    
    Args:
        filepath: Path to Excel file
    
    Returns:
        FullQuestionBank with all questions loaded
    """
    # Read Excel file
    df = pd.read_excel(filepath)
    
    # Normalize column names (lowercase, strip whitespace)
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
    
    # Validate required columns
    required_cols = ['question_no', 'question', 'option_a', 'option_b', 
                     'option_c', 'option_d', 'answer', 'difficulty']
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    
    questions = []
    counters = {'hard': 0, 'medium': 0, 'easy': 0}
    
    for _, row in df.iterrows():
        difficulty = normalize_difficulty(row['difficulty'])
        counters[difficulty] += 1
        
        # Generate internal ID (H1, H2, M1, M2, E1, E2, etc.)
        prefix = {'hard': 'H', 'medium': 'M', 'easy': 'E'}[difficulty]
        question_id = f"{prefix}{counters[difficulty]}"
        
        q = FullQuestion(
            question_no=int(row['question_no']),
            question_id=question_id,
            question_text=str(row['question']),
            option_a=str(row['option_a']),
            option_b=str(row['option_b']),
            option_c=str(row['option_c']),
            option_d=str(row['option_d']),
            answer=str(row['answer']).strip().upper(),
            difficulty=difficulty
        )
        questions.append(q)
    
    return FullQuestionBank(questions)


def generate_question_papers(
    allocation_matrix: List[List[str]],
    question_bank: FullQuestionBank,
    output_path: str,
    include_answer_key: bool = True
) -> str:
    """
    Generate question papers as multi-sheet Excel file.
    
    Each student gets a separate sheet named "Set_1", "Set_2", etc.
    Questions are shown WITHOUT the answer column.
    
    Args:
        allocation_matrix: 2D list [num_students][num_questions] of question_ids
        question_bank: FullQuestionBank with complete question data
        output_path: Path for output Excel file
        include_answer_key: If True, add an answer key sheet for teachers
    
    Returns:
        Path to created file
    """
    # Ensure output directory exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Generate sheet for each student
        for student_idx, quiz in enumerate(allocation_matrix):
            sheet_name = f"Set_{student_idx + 1}"
            
            rows = []
            for q_idx, question_id in enumerate(quiz):
                q = question_bank.get_by_id(question_id)
                if q is None:
                    raise ValueError(f"Question ID not found: {question_id}")
                
                rows.append({
                    'Q.No': q_idx + 1,  # Sequential numbering in quiz
                    'Question': q.question_text,
                    'Option A': q.option_a,
                    'Option B': q.option_b,
                    'Option C': q.option_c,
                    'Option D': q.option_d,
                })
            
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Optional: Add answer key sheet for teachers
        if include_answer_key:
            answer_rows = []
            for student_idx, quiz in enumerate(allocation_matrix):
                student_answers = {}
                student_answers['Set'] = f"Set_{student_idx + 1}"
                for q_idx, question_id in enumerate(quiz):
                    q = question_bank.get_by_id(question_id)
                    student_answers[f'Q{q_idx + 1}'] = q.answer
                answer_rows.append(student_answers)
            
            answer_df = pd.DataFrame(answer_rows)
            answer_df.to_excel(writer, sheet_name='Answer_Key', index=False)
    
    return output_path


def create_sample_question_bank_excel(
    filepath: str,
    hard_count: int = 10,
    medium_count: int = 25,
    easy_count: int = 15
) -> str:
    """
    Create a sample question bank Excel file for testing.
    
    Generates placeholder questions with the correct format.
    """
    rows = []
    q_no = 1
    
    # Hard questions
    for i in range(1, hard_count + 1):
        rows.append({
            'question_no': q_no,
            'question': f"Hard Question {i}: What is the solution to this complex problem?",
            'option_a': f"Hard option A{i}",
            'option_b': f"Hard option B{i}",
            'option_c': f"Hard option C{i}",
            'option_d': f"Hard option D{i}",
            'answer': ['A', 'B', 'C', 'D'][i % 4],
            'difficulty': 'H'
        })
        q_no += 1
    
    # Medium questions
    for i in range(1, medium_count + 1):
        rows.append({
            'question_no': q_no,
            'question': f"Medium Question {i}: Calculate the following expression.",
            'option_a': f"Medium option A{i}",
            'option_b': f"Medium option B{i}",
            'option_c': f"Medium option C{i}",
            'option_d': f"Medium option D{i}",
            'answer': ['A', 'B', 'C', 'D'][i % 4],
            'difficulty': 'M'
        })
        q_no += 1
    
    # Easy questions
    for i in range(1, easy_count + 1):
        rows.append({
            'question_no': q_no,
            'question': f"Easy Question {i}: What is the basic definition of this term?",
            'option_a': f"Easy option A{i}",
            'option_b': f"Easy option B{i}",
            'option_c': f"Easy option C{i}",
            'option_d': f"Easy option D{i}",
            'answer': ['A', 'B', 'C', 'D'][i % 4],
            'difficulty': 'L'
        })
        q_no += 1
    
    df = pd.DataFrame(rows)
    df.to_excel(filepath, index=False)
    
    return filepath
