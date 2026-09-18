@app.post("/api/commitments")
def commitments(x: CommitmentReq):
    try:
        # Keep enough context for semantic matching while limiting provider payload size.
        def compact_mail(m, direction):
            body = ((m.get("body") or {}).get("content") if isinstance(m.get("body"), dict) else None) or m.get("bodyPreview") or ""
            return {
                "id": m.get("id", ""),
                "direction": direction,
                "subject": m.get("subject", ""),
                "from": m.get("from", {}),
                "toRecipients": m.get("toRecipients", []),
                "ccRecipients": m.get("ccRecipients", []),
                "timestamp": m.get("timestamp") or m.get("sentDateTime") or m.get("receivedDateTime") or "",
                "body": str(body)[:3500],
                "webLink": m.get("webLink", ""),
                "conversationId": m.get("conversationId", "")
            }
        sent = [compact_mail(m, "sent") for m in x.sent[:30]]
        inbox = [compact_mail(m, "received") for m in x.inbox[:30]]
        user_content = {
            "NOW": x.now or datetime.now(timezone.utc).isoformat(),
            "EMPLOYEE": x.employee,
            "SENT_EMAILS": sent,
            "INBOX_EMAILS": inbox
        }
        content = provider_call([
            {"role": "system", "content": COMMITMENT_SYSTEM},
            {"role": "user", "content": json.dumps(user_content, ensure_ascii=False)}
        ], temperature=0.05)
        data = json.loads(content)
        data.setdefault("generatedAt", datetime.now(timezone.utc).isoformat())
        data.setdefault("summary", {})
        data.setdefault("commitments", [])
        return data
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Commitment AI returned invalid JSON: {e}")
    except Exception as e:
        raise HTTPException(500, str(e))


class DraftReq(BaseModel):
    commitment: dict = Field(default_factory=dict)

DRAFT_SYSTEM = """
You draft concise professional workplace follow-up emails.
Use only the supplied commitment. Do not invent project facts or deadlines.
If a recipient email is present, use it. If not, return an empty `to`.
Return JSON ONLY: {"to":"","subject":"","body":""}.
The message should politely reference the commitment and ask for/communicate the next step.
If the commitment is overdue, acknowledge the follow-up without inventing an excuse.
"""

@app.post("/api/commitments/draft")
def commitment_draft(x: DraftReq):
    try:
        content = provider_call([
            {"role": "system", "content": DRAFT_SYSTEM},
            {"role": "user", "content": json.dumps(x.commitment, ensure_ascii=False)}
        ], temperature=0.2)
        data = json.loads(content)
        return {
            "to": data.get("to", ""),
            "subject": data.get("subject", f"Follow-up: {x.commitment.get('title', 'Commitment')}"),
            "body": data.get("body", "")
        }
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Draft AI returned invalid JSON: {e}")
    except Exception as e:
        raise HTTPException(500, str(e))
