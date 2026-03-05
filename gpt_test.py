import json
import urllib.request
import urllib.error


class GPT_test:
    """
    Optional LLM-backed grader.
    - Can run fully offline (heuristic fallback) when key is missing.
    - When API key exists, sends one question at a time with code + rubric.
    """

    def __init__(self, api_key: str = "", model: str = "gpt-4.1-mini", system_prompt: str = ""):
        self.api_key = (api_key or "").strip()
        self.model = (model or "gpt-4.1-mini").strip()
        self.system_prompt = (system_prompt or "").strip()
        self.last_prompt_payload = None
        self.last_request_body = None
        self.last_raw_response = None
        self.last_result = None

    def _heuristic_grade(self, question_id: str, question_title: str, rubric_items: list[dict], code_text: str) -> dict:
        text = (code_text or "").lower()
        scores = []
        lines = [ln for ln in (code_text or "").splitlines() if ln.strip()]
        line_count = len(lines)
        comment_ratio = 0.0
        if lines:
            c_lines = sum(1 for ln in lines if ln.strip().startswith(("//", "/*", "*")))
            comment_ratio = c_lines / max(1, line_count)

        for item in (rubric_items or []):
            col_key = item.get("col_key")
            if not col_key:
                continue
            mx = float(item.get("max_points", 0.0) or 0.0)
            crit = (item.get("criterion") or "").lower()

            ratio = 0.55
            if any(k in crit for k in ["output", "result", "correct"]):
                ratio = 0.85 if ("println" in text or "return" in text) else 0.4
            elif any(k in crit for k in ["style", "naming", "format"]):
                ratio = min(1.0, 0.45 + comment_ratio)
            elif any(k in crit for k in ["loop", "iteration"]):
                ratio = 0.9 if ("for(" in text or "while(" in text or "for (" in text or "while (" in text) else 0.35
            elif any(k in crit for k in ["condition", "if", "branch"]):
                ratio = 0.9 if "if(" in text or "if (" in text else 0.35

            points = round(max(0.0, min(mx, mx * ratio)), 2)
            note = f"Estimated against criterion text for {question_id}."
            scores.append({"col_key": col_key, "points": points, "note": note})

        rationale = (
            f"Feedback summary for {question_id} ({question_title or question_id}) based on rubric criteria and code evidence."
        )
        comments = []
        for idx, ln in enumerate((code_text or "").splitlines(), start=1):
            low = ln.lower()
            if "todo" in low or "fixme" in low:
                comments.append({"line": idx, "comment": "TODO/FIXME found: complete this logic before final submission."})
            if "system.out.print" in low and "println" not in low:
                comments.append({"line": idx, "comment": "Consider clearer output formatting with println."})
            if len(comments) >= 6:
                break

        return {"scores": scores, "rationale": rationale, "comments": comments}

    def grade_question(self, *, question_id: str, question_title: str = "", rubric_items: list[dict], code_text: str, extra_prompt: str = "") -> dict:
        self.last_raw_response = None
        self.last_result = None
        if not self.api_key:
            self.last_prompt_payload = {
                "question_id": question_id,
                "question_title": (question_title or question_id),
                "mode": "heuristic_fallback",
                "rubric_items": rubric_items or [],
                "code": code_text,
                "extra_prompt": extra_prompt or "",
            }
            self.last_request_body = None
            result = self._heuristic_grade(question_id, question_title, rubric_items, code_text)
            self.last_result = result
            return result

        prompt = {
            "question_id": question_id,
            "question_title": (question_title or question_id),
            "instructions": (self.system_prompt or "Grade code strictly but fairly.") + "\n" + (extra_prompt or "") + "\nUse the exact question context and rubric min/max ranges.",
            "rubric_items": [{**ri, "min_points": float(ri.get("min_points", 0.0) or 0.0), "max_points": float(ri.get("max_points", 0.0) or 0.0)} for ri in (rubric_items or [])],
            "code": code_text,
            "output_format": {
                "scores": [{"col_key": "string", "points": "number", "note": "string"}],
                "rationale": "string",
                "comments": [{"line": "number", "comment": "string"}],
            },
        }
        req_body = json.dumps({
            "model": self.model,
            "input": [
                {"role": "system", "content": [{"type": "input_text", "text": "Return strict JSON only."}]},
                {"role": "user", "content": [{"type": "input_text", "text": json.dumps(prompt)}]},
            ],
        }).encode("utf-8")
        self.last_prompt_payload = prompt
        self.last_request_body = req_body.decode("utf-8", errors="ignore")

        req = urllib.request.Request(
            "https://api.openai.com/v1/responses",
            data=req_body,
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=45) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"GPT request failed: {e.code} {detail[:300]}")
        except Exception as e:
            raise RuntimeError(f"GPT request failed: {e}")

        self.last_raw_response = payload
        text = payload.get("output_text", "").strip()
        if not text:
            try:
                text = payload["output"][0]["content"][0]["text"]
            except Exception:
                text = ""
        try:
            parsed = json.loads(text)
        except Exception as e:
            raise RuntimeError(f"GPT returned invalid JSON: {e}")
        self.last_result = parsed
        return parsed

    def chat(self, *, message: str, context_bundle: str = "") -> str:
        """
        Lightweight chat helper for the UI.
        - Offline mode: returns a concise, deterministic helper response.
        - Online mode: calls the Responses API using the configured model.
        """
        clean_msg = (message or "").strip()
        clean_ctx = (context_bundle or "").strip()
        if not clean_msg:
            return "Please enter a message first."

        if not self.api_key:
            lines = [
                "Local assistant mode is active.",
                f"You said: {clean_msg[:300]}",
            ]
            if clean_ctx:
                snippet = "\n".join(clean_ctx.splitlines()[:8]).strip()
                ctx_lines = len(clean_ctx.splitlines())
                lines.append(f"Bundle included: {ctx_lines} line(s).")
                lines.append("Quick local analysis (no API key):")
                if snippet:
                    lines.append(snippet[:500])
                lines.append("I can summarize/reformat this bundle locally. Model-based scoring still needs an API key.")
            else:
                lines.append("Add a bundle on the left panel for richer context.")
            return "\n".join(lines)

        user_payload = {
            "message": clean_msg,
            "context_bundle": clean_ctx,
            "instructions": "Answer clearly and concisely. If context_bundle is provided, use it.",
        }
        req_body = json.dumps({
            "model": self.model,
            "input": [
                {"role": "system", "content": [{"type": "input_text", "text": self.system_prompt or "You are a grading assistant."}]},
                {"role": "user", "content": [{"type": "input_text", "text": json.dumps(user_payload)}]},
            ],
        }).encode("utf-8")

        req = urllib.request.Request(
            "https://api.openai.com/v1/responses",
            data=req_body,
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=45) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"Chat request failed: {e.code} {detail[:300]}")
        except Exception as e:
            raise RuntimeError(f"Chat request failed: {e}")

        text = payload.get("output_text", "").strip()
        if text:
            return text
        try:
            return payload["output"][0]["content"][0]["text"].strip()
        except Exception:
            return "Model returned no text."
