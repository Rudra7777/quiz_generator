"""
Quiz Question Allocator Module

Implements a greedy load-balancing algorithm using min-heap priority queues
to ensure uniform question usage across all students.

Supports dynamic configuration:
- Any number of questions in the bank
- Any number of students
- Any quiz structure (configurable questions per difficulty)
"""

import heapq
import random
from dataclasses import dataclass, field
from typing import Dict, List, Set, Tuple, Optional
from collections import defaultdict


@dataclass
class Question:
    """Represents a single question with its properties."""
    question_id: str
    difficulty: str  # 'hard', 'medium', 'easy'


@dataclass
class QuestionBank:
    """
    Stores questions grouped by difficulty.
    Provides methods to retrieve questions by difficulty level.
    """
    questions: List[Question] = field(default_factory=list)
    
    def __post_init__(self):
        self._by_difficulty: Dict[str, List[Question]] = defaultdict(list)
        for q in self.questions:
            self._by_difficulty[q.difficulty].append(q)
    
    def get_by_difficulty(self, difficulty: str) -> List[Question]:
        """Get all questions of a specific difficulty."""
        return self._by_difficulty[difficulty]
    
    def get_all(self) -> List[Question]:
        """Get all questions in the bank."""
        return self.questions
    
    def count_by_difficulty(self) -> Dict[str, int]:
        """Get count of questions per difficulty level."""
        return {d: len(qs) for d, qs in self._by_difficulty.items()}
    
    def get_question_ids_by_difficulty(self, difficulty: str) -> List[str]:
        """Get list of question IDs for a difficulty level."""
        return [q.question_id for q in self._by_difficulty.get(difficulty, [])]


class UsageTracker:
    """
    Tracks usage counts for all questions using min-heaps.
    
    Maintains separate heaps per difficulty for efficient retrieval
    of the least-used eligible question.
    """
    
    def __init__(self, question_ids_by_difficulty: Dict[str, List[str]]):
        """
        Initialize tracker with question IDs grouped by difficulty.
        
        Args:
            question_ids_by_difficulty: Dict mapping difficulty to list of question_ids
        """
        # Primary usage count storage: question_id -> count
        self.usage_counts: Dict[str, int] = {}
        # Min-heaps per difficulty: (count, tiebreaker, question_id)
        self._heaps: Dict[str, List[Tuple[int, float, str]]] = defaultdict(list)
        
        # Initialize all questions with 0 usage
        for difficulty, q_ids in question_ids_by_difficulty.items():
            for qid in q_ids:
                self.usage_counts[qid] = 0
                heapq.heappush(
                    self._heaps[difficulty],
                    (0, random.random(), qid)
                )
    
    def get_least_used(self, difficulty: str, excluded: Set[str]) -> str:
        """
        Get the question_id with minimum usage count for given difficulty,
        excluding questions already assigned to the current student.
        """
        heap = self._heaps[difficulty]
        candidates = []
        
        while heap:
            count, tiebreaker, qid = heapq.heappop(heap)
            actual_count = self.usage_counts[qid]
            
            if count != actual_count:
                heapq.heappush(heap, (actual_count, random.random(), qid))
            elif qid in excluded:
                candidates.append((count, random.random(), qid))
            else:
                for c in candidates:
                    heapq.heappush(heap, c)
                self._pending_push = (difficulty, qid)
                return qid
        
        for c in candidates:
            heapq.heappush(heap, c)
        
        raise ValueError(
            f"No eligible questions available for difficulty '{difficulty}'. "
            f"Excluded: {len(excluded)}"
        )
    
    def increment_usage(self, question_id: str):
        """Increment the usage count for a question and update the heap."""
        self.usage_counts[question_id] += 1
        new_count = self.usage_counts[question_id]
        
        if hasattr(self, '_pending_push'):
            difficulty, qid = self._pending_push
            if qid == question_id:
                heapq.heappush(
                    self._heaps[difficulty],
                    (new_count, random.random(), qid)
                )
            delattr(self, '_pending_push')
    
    def get_usage_counts(self) -> Dict[str, int]:
        """Get a copy of all usage counts."""
        return dict(self.usage_counts)


@dataclass
class QuizStructure:
    """
    Defines the structure of a quiz with questions per difficulty.
    """
    hard_count: int = 4
    medium_count: int = 6
    easy_count: int = 5
    
    def get_structure(self) -> List[Tuple[str, int]]:
        """
        Returns list of (difficulty, count) tuples.
        """
        return [
            ('hard', self.hard_count),
            ('medium', self.medium_count),
            ('easy', self.easy_count),
        ]
    
    def total_questions(self) -> int:
        return self.hard_count + self.medium_count + self.easy_count
    
    def validate(self, question_counts: Dict[str, int]) -> Tuple[bool, List[str]]:
        """
        Validate that question bank has enough questions for this structure.
        """
        errors = []
        for difficulty, needed in self.get_structure():
            available = question_counts.get(difficulty, 0)
            if available < needed:
                errors.append(
                    f"Need {needed} {difficulty} questions, only {available} available"
                )
        return (len(errors) == 0, errors)


def allocate_quizzes(
    question_ids_by_difficulty: Dict[str, List[str]],
    num_students: int,
    quiz_structure: QuizStructure,
    seed: int = None
) -> Tuple[List[List[str]], Dict[str, int]]:
    """
    Allocate questions to students using greedy load-balancing.
    
    Args:
        question_ids_by_difficulty: Dict mapping difficulty to list of question_ids
        num_students: Number of students to generate quizzes for
        quiz_structure: QuizStructure defining H/M/E counts
        seed: Random seed for reproducibility
    
    Returns:
        Tuple of:
        - allocation_matrix: 2D list [num_students][total_questions] of question_ids
        - usage_counts: Dict mapping question_id to total usage count
    """
    if seed is not None:
        random.seed(seed)
    
    # Validate
    counts = {d: len(ids) for d, ids in question_ids_by_difficulty.items()}
    valid, errors = quiz_structure.validate(counts)
    if not valid:
        raise ValueError(f"Invalid configuration: {errors}")
    
    # Initialize tracker
    tracker = UsageTracker(question_ids_by_difficulty)
    
    # Allocation matrix
    allocation_matrix: List[List[str]] = []
    
    for student_idx in range(num_students):
        student_quiz: List[str] = []
        assigned_to_student: Set[str] = set()
        
        for difficulty, count in quiz_structure.get_structure():
            for _ in range(count):
                qid = tracker.get_least_used(difficulty, assigned_to_student)
                student_quiz.append(qid)
                assigned_to_student.add(qid)
                tracker.increment_usage(qid)
        
        allocation_matrix.append(student_quiz)
    
    return allocation_matrix, tracker.get_usage_counts()


def shuffle_quiz(quiz: list, seed: int = None) -> list:
    """Shuffle questions within a quiz."""
    shuffled = quiz.copy()
    if seed is not None:
        random.seed(seed)
    random.shuffle(shuffled)
    return shuffled


def shuffle_all_quizzes(allocation_matrix: list, base_seed: int = None) -> list:
    """Shuffle questions for all students."""
    shuffled_matrix = []
    for i, quiz in enumerate(allocation_matrix):
        seed = base_seed + i if base_seed is not None else None
        shuffled_matrix.append(shuffle_quiz(quiz, seed))
    return shuffled_matrix


def create_sample_question_bank(
    hard_count: int = 10,
    medium_count: int = 25,
    easy_count: int = 15
) -> QuestionBank:
    """
    Create a sample question bank with specified counts per difficulty.
    """
    questions = []
    
    for i in range(1, hard_count + 1):
        questions.append(Question(f"H{i}", "hard"))
    
    for i in range(1, medium_count + 1):
        questions.append(Question(f"M{i}", "medium"))
    
    for i in range(1, easy_count + 1):
        questions.append(Question(f"E{i}", "easy"))
    
    return QuestionBank(questions)
