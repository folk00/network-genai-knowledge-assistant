# Architecture Diagrams

These diagrams are safe for GitHub because they describe the system architecture
without showing private question banks, customer data or generated reports.

## System Overview

```mermaid
flowchart LR
    User[User / Network Engineer] --> GUI[PySide6 Desktop GUI]

    GUI --> State[Deterministic State Layer<br/>profiles, wrong bank, confidence, stats]
    GUI --> Workflows[AI Workflow Buttons<br/>AI Coach, Deep Review, Nuclear, Diagram, Teach Zero]
    GUI --> RAG[RAG KB<br/>grounded context retrieval]

    State --> Reports[Local Markdown / HTML Reports]
    Workflows --> Provider[Provider Adapter<br/>Claude / OpenAI style backends]
    RAG --> Retriever[Retrieval Layer<br/>chunks, scores, citations]
    Retriever --> Prompt[Grounded Prompt Builder]
    Prompt --> Provider
    Provider --> Reports

    Reports --> Reuse[Reusable Context<br/>future review, dossiers, meta-coach]
    Reuse --> Workflows
```

## RAG KB Flow In The Full GUI

```mermaid
sequenceDiagram
    actor User
    participant GUI as Full PySide6 GUI
    participant Bank as Loaded Question Bank
    participant RAG as RAG Retriever
    participant LLM as Selected AI Backend
    participant Report as Report Dialog

    User->>GUI: Click RAG KB / Grounded Context
    GUI->>GUI: Read current question and selected answer
    GUI->>Bank: Build temporary source docs from loaded bank
    GUI->>Bank: Exclude current question
    GUI->>RAG: Retrieve related chunks
    RAG-->>GUI: Context chunks + scores + citations
    GUI->>GUI: Add optional private/mock knowledge docs
    GUI->>User: Confirm before sending context
    User-->>GUI: Approve
    GUI->>LLM: Send grounded study prompt
    LLM-->>GUI: Study review with citations
    GUI->>Report: Render Markdown dialog with prompt details
```

## Study Workflow To Network Operations Mapping

```mermaid
flowchart TB
    subgraph Study_Workflow[Current Study Workflow]
        Q[Question + wrong answer] --> C[Concept / tags / pattern]
        C --> A[AI Coach or Deep Review]
        A --> SR[Study report]
    end

    subgraph Network_Ops_Workflow[Equivalent Network Operations Workflow]
        I[Incident / config / log / MOP] --> K[Runbook + topology + prior cases]
        K --> D[Diagnostic workflow]
        D --> OR[Operations report<br/>validation + rollback + next actions]
    end

    SR -.same architecture pattern.-> OR
```

## Data Safety Boundary

```mermaid
flowchart LR
    subgraph Public_Repo[Public GitHub Repo]
        Code[Application Code]
        Mock[Synthetic Mock Data]
        Docs[Architecture Docs]
        Tests[Tests + Safety Scan]
    end

    subgraph Local_Private[Local Private Machine]
        Banks[Real DOCX / CSV Inputs]
        Env[.env / API Keys]
        Logs[Logs]
        Generated[Generated Reports / Dossiers]
    end

    Code --> Runtime[Runtime Only]
    Runtime --> Banks
    Runtime --> Env
    Runtime --> Generated

    Local_Private -.blocked by .gitignore.-> Public_Repo
```

## Provider And Safety Model

```mermaid
flowchart TD
    Prompt[Controlled Prompt] --> Confirm{User Confirmation?}
    Confirm -- no --> Stop[Do Not Call Provider]
    Confirm -- yes --> Provider{Selected Backend}

    Provider --> Claude[Claude Code / Claude-compatible path]
    Provider --> OpenAI[OpenAI-compatible path]

    Claude --> Output[Markdown / HTML Artifact]
    OpenAI --> Output

    Output --> Human[Human Review]
    Human --> Persist[Persist Local Report]
```

## Future Enterprise Version

```mermaid
flowchart TB
    Sources[Configs, MOPs, Logs, Runbooks, Diagrams] --> Ingest[Ingestion + Redaction]
    Ingest --> Chunks[Chunking + Metadata]
    Chunks --> Index[Hybrid Index<br/>keyword + vector]
    Index --> Filters[Tenant / site / source filters]
    Filters --> Context[Retrieved Context + Citations]
    Context --> Agent[Controlled Agent Workflow]
    Agent --> Draft[Diagnosis / Validation / Rollback Draft]
    Draft --> Approval[Human Approval]
    Approval --> Audit[Audit Trail + Report Store]
```
