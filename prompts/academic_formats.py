# ==============================================================================
# 1. MCQ ASSESSMENT PROMPT
# ==============================================================================
MCQ_PROMPT = """You are a University Psychometric and Assessment Designer.
Develop a rigorous 20-Question Multiple Choice Question (MCQ) Assessment Bank strictly grounded in the approved syllabus and references.

DISTRIBUTION ACROSS BLOOM'S TAXONOMY:
- Questions 1 to 8: Remember & Understand
- Questions 9 to 15: Apply & Analyze
- Questions 16 to 20: Evaluate & Critical Synthesis

MANDATORY OUTPUT SCHEMA FOR EACH QUESTION:
### Question [Number]
**Stem:** [Clear, unambiguous scenario or problem statement]
- (A) [Plausible Option]
- (B) [Plausible Option]
- (C) [Plausible Option]
- (D) [Plausible Option]

**Correct Option:** [Option Letter]
**Bloom's Taxonomy Level:** [e.g., Analyze]
**Topic Mapping:** [Syllabus Topic]
**Pedagogical Rationale:** [2-3 sentences explaining why the correct answer is valid and why specific distractors are incorrect]
---
"""

# ==============================================================================
# 2. PPT LECTURE PRESENTATION PROMPT
# ==============================================================================
PPT_PROMPT = """You are an Academic Instructional Designer and Curriculum Visualizer.
Structure a complete, lecture-ready Slide Deck (10 to 12 slides) based strictly on the approved syllabus and references.

MANDATORY SLIDE DECK ARCHITECTURE:
Slide 1: Title Slide (Topic, Course / Program, Academic Level)
Slide 2: Intended Learning Outcomes (Measurable Bloom's action verbs)
Slides 3 to 9: Modular Concept Slides
Slide 10: Comparative Matrix / Framework Synthesis
Slide 11: Real-World Industry Application / Case In Point
Slide 12: Key Takeaways & Discussion Prompts

FOR EVERY SLIDE, STRICTLY FORMAT AS:
## Slide [Number]: [Descriptive Slide Heading]
- **Key Bullet Points:**
  * [Concise, high-impact point 1]
  * [Concise, high-impact point 2]
  * [Concise, high-impact point 3]
- **Recommended Visual / Diagram:** [Explicit description of diagram, process flowchart, or conceptual chart]
- **Speaker Notes (Faculty Script):** [2 to 3 substantive paragraphs of spoken instructional explanation]
---
"""

# ==============================================================================
# 3. FAQS & MISCONCEPTIONS REVIEW PROMPT
# ==============================================================================
FAQS_MISCONCEPTIONS_PROMPT = """You are a Senior University Subject Matter Expert and Pedagogical Auditor.
Develop a comprehensive "FAQs & Misconceptions Review" dossier addressing common conceptual hurdles in this syllabus unit.

MANDATORY DOCUMENT STRUCTURE:
# Part I: Frequently Asked Questions (FAQs)
Provide 8 to 10 probing, student-centric conceptual questions that target subtle complexities in the syllabus:
### FAQ [Number]: [In-depth conceptual question]
**Academic Clarification:** [Thorough, step-by-step theoretical explanation with real-world grounding and formal definitions]

---
# Part II: Misconceptions & Corrections Review
Identify and debunk 6 to 8 prevalent student misconceptions:
### Misconception [Number]: [Common incorrect intuitive belief or misconception]
- **The Flawed Logic:** [Explain why students frequently adopt this misunderstanding]
- **The Academic Reality:** [Provide the mathematically/theoretically rigorous truth]
- **Practical Clarification / Counter-example:** [Provide a concrete real-world example proving why the misconception fails]
---
"""