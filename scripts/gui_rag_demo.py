from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSplitter,
    QStatusBar,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from network_ai_assistant.llm.clients import build_client
from network_ai_assistant.llm.safety import load_env_file
from network_ai_assistant.rag.ingest import chunk_documents, load_markdown_documents
from network_ai_assistant.rag.prompting import build_grounded_prompt, deterministic_summary
from network_ai_assistant.rag.simple_retriever import KeywordRetriever


DEFAULT_QUESTION = (
    "A Cisco branch advertises 172.16.20.0/24 to AWS through a Site-to-Site VPN "
    "attached to Transit Gateway. What should I validate if the VPC cannot reach "
    "the branch LAN?"
)


class AnswerWorker(QObject):
    finished = Signal(str, str)

    def __init__(self, prompt: str, provider: str, live: bool):
        super().__init__()
        self.prompt = prompt
        self.provider = provider
        self.live = live

    @Slot()
    def run(self) -> None:
        previous_live = os.environ.get("AI_LIVE")
        previous_provider = os.environ.get("AI_PROVIDER")
        try:
            os.environ["AI_LIVE"] = "1" if self.live else "0"
            os.environ["AI_PROVIDER"] = self.provider
            client = build_client(self.provider)
            self.finished.emit(client.generate(self.prompt), "")
        except Exception as exc:
            self.finished.emit("", f"{type(exc).__name__}: {exc}")
        finally:
            if previous_live is None:
                os.environ.pop("AI_LIVE", None)
            else:
                os.environ["AI_LIVE"] = previous_live
            if previous_provider is None:
                os.environ.pop("AI_PROVIDER", None)
            else:
                os.environ["AI_PROVIDER"] = previous_provider


class RAGDemoWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Network AI Knowledge Assistant - RAG Demo")
        self.resize(1180, 760)

        load_env_file(ROOT / ".env")

        self._thread: QThread | None = None
        self._worker: AnswerWorker | None = None

        self.docs_path = ROOT / "data" / "mock" / "network_docs.md"
        self.retrieved_prompt = ""

        self._build_ui()
        self._load_sources()

    def _build_ui(self) -> None:
        central = QWidget(self)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        top = QFormLayout()
        self.source_path = QLineEdit(str(self.docs_path))
        self.source_path.setReadOnly(True)
        browse = QPushButton("Browse")
        browse.clicked.connect(self._browse_source)
        source_row = QHBoxLayout()
        source_row.addWidget(self.source_path, 1)
        source_row.addWidget(browse)
        top.addRow("Knowledge source", source_row)

        self.provider = QComboBox()
        self.provider.addItems(["claude_code", "openai", "anthropic"])
        self.provider.setCurrentText("claude_code")
        self.live_calls = QCheckBox("Live LLM call")
        self.live_calls.setToolTip("Unchecked = dry-run. Checked = send mock prompt to the selected provider.")
        provider_row = QHBoxLayout()
        provider_row.addWidget(self.provider)
        provider_row.addWidget(self.live_calls)
        provider_row.addStretch(1)
        top.addRow("Provider", provider_row)
        root.addLayout(top)

        self.question = QPlainTextEdit()
        self.question.setPlainText(DEFAULT_QUESTION)
        self.question.setMinimumHeight(90)
        root.addWidget(QLabel("Question"))
        root.addWidget(self.question)

        buttons = QHBoxLayout()
        self.btn_retrieve = QPushButton("Retrieve Context")
        self.btn_retrieve.clicked.connect(self.retrieve_context)
        self.btn_ask = QPushButton("Ask LLM")
        self.btn_ask.clicked.connect(self.ask_llm)
        self.btn_scan = QPushButton("Safety Scan")
        self.btn_scan.clicked.connect(self.run_safety_scan)
        buttons.addWidget(self.btn_retrieve)
        buttons.addWidget(self.btn_ask)
        buttons.addWidget(self.btn_scan)
        buttons.addStretch(1)
        root.addLayout(buttons)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.tabs_left = QTabWidget()
        self.context_view = QPlainTextEdit()
        self.context_view.setReadOnly(True)
        self.prompt_view = QPlainTextEdit()
        self.prompt_view.setReadOnly(True)
        self.tabs_left.addTab(self.context_view, "Retrieved Context")
        self.tabs_left.addTab(self.prompt_view, "Grounded Prompt")

        self.answer_view = QPlainTextEdit()
        self.answer_view.setReadOnly(True)
        self.answer_view.setPlaceholderText("Run dry-run or live Claude Code to see the grounded answer.")

        splitter.addWidget(self.tabs_left)
        splitter.addWidget(self.answer_view)
        splitter.setSizes([560, 620])
        root.addWidget(splitter, 1)

        self.setCentralWidget(central)
        self.status = QStatusBar(self)
        self.setStatusBar(self.status)

    def _load_sources(self) -> None:
        try:
            docs = load_markdown_documents(self.docs_path)
            chunks = chunk_documents(docs)
            self.retriever = KeywordRetriever(chunks)
            self.status.showMessage(f"Loaded {len(docs)} docs / {len(chunks)} chunks from {self.docs_path.name}")
        except Exception as exc:
            QMessageBox.critical(self, "Load failed", str(exc))

    def _browse_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Pick sanitized markdown knowledge source",
            str(ROOT / "data"),
            "Markdown files (*.md);;Text files (*.txt);;All files (*.*)",
        )
        if not path:
            return
        self.docs_path = Path(path)
        self.source_path.setText(str(self.docs_path))
        self._load_sources()

    def retrieve_context(self) -> None:
        q = self.question.toPlainText().strip()
        if not q:
            QMessageBox.information(self, "Question missing", "Type a question first.")
            return
        results = self.retriever.search(q, k=3, filters={"tenant": "demo-retail"})
        self.retrieved_prompt = build_grounded_prompt(q, results)
        self.context_view.setPlainText(deterministic_summary(q, results))
        self.prompt_view.setPlainText(self.retrieved_prompt)
        self.tabs_left.setCurrentWidget(self.context_view)
        self.status.showMessage(f"Retrieved {len(results)} context chunks.")

    def ask_llm(self) -> None:
        if not self.retrieved_prompt:
            self.retrieve_context()
        if not self.retrieved_prompt:
            return

        live = self.live_calls.isChecked()
        provider = self.provider.currentText()
        if live:
            ans = QMessageBox.question(
                self,
                "Confirm live call",
                "This will send the sanitized mock prompt to the selected provider.\n\n"
                "Do not use this with real quiz banks or customer data.\n\n"
                f"Provider: {provider}\nProceed?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel,
            )
            if ans != QMessageBox.StandardButton.Yes:
                return

        self.btn_ask.setEnabled(False)
        self.answer_view.setPlainText("Running...")
        self.status.showMessage(f"Calling {provider} ({'live' if live else 'dry-run'})...")

        self._thread = QThread(self)
        self._worker = AnswerWorker(self.retrieved_prompt, provider, live)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._answer_finished)
        self._thread.start()

    def _answer_finished(self, answer: str, error: str) -> None:
        worker = self._worker
        thread = self._thread
        self._worker = None
        self._thread = None
        if worker:
            worker.deleteLater()
        if thread:
            thread.quit()
            thread.deleteLater()
        self.btn_ask.setEnabled(True)

        if error:
            self.answer_view.setPlainText(error)
            self.status.showMessage("LLM call failed.")
            return
        self.answer_view.setPlainText(answer)
        self.status.showMessage("Answer generated.")

    def run_safety_scan(self) -> None:
        proc = subprocess.run(
            [sys.executable, str(ROOT / "scripts" / "safety_scan.py")],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
        )
        output = (proc.stdout or "") + (proc.stderr or "")
        self.answer_view.setPlainText(output.strip())
        self.status.showMessage("Safety scan passed." if proc.returncode == 0 else "Safety scan found issues.")


def main() -> int:
    if "--smoke" in sys.argv:
        print("gui_rag_demo import smoke OK")
        return 0
    app = QApplication(sys.argv)
    window = RAGDemoWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
