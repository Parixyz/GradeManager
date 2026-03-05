from gpt_test import GPT_test


class AutoGrader:
    def __init__(self, gpt_client: GPT_test | None = None):
        self.gpt_client = gpt_client or GPT_test()

    def auto_grade(self, *, question_id: str, merged_code: str, rubric_items: list[dict], theme_text: str) -> dict:
        return self.gpt_client.grade_question(
            question_id=question_id,
            rubric_items=rubric_items,
            code_text=merged_code,
            extra_prompt=theme_text,
        )
