from pathlib import Path
import colorlog
from datetime import datetime
import logging
import os

PROJ_ABSPATH = Path(__file__).resolve().parent.absolute()


def trancate(text: str, max_length: int = 35, placeholder: str = "...") -> str:
    if len(text) <= max_length:
        return text
    return text[: max_length - len(placeholder)] + placeholder


# 1. シーン: プロジェクト全体のロギング設定を簡単に管理したいときに利用するクラス
# 2. 呼び元: main関数やwith構文などから直接呼ばれることを想定
# 3. 内容: ログファイル・コンソール出力・色付け・フィルタなどを一括で制御する
class SmartLogging:
    # 1. シーン: SmartLoggingインスタンス生成時に呼ばれる初期化メソッド
    # 2. 呼び元: SmartLogging()の呼び出し全般で利用されることを想定
    # 3. 内容: ログレベルやストリームハンドラの初期値を設定する
    def __init__(self, level=logging.INFO) -> None:
        self._level = level
        self._stream_handler: logging.StreamHandler | None = None

    # 1. シーン: with構文でロギング環境を自動セットアップしたいときに使用
    # 2. 呼び元: with SmartLogging(...) as env: などから呼ばれることを想定
    # 3. 内容: ロギング初期化を行い、コンテキストマネージャとして自身を返す
    def __enter__(self):
        self.initialize_logging()
        return self

    # 1. シーン: with構文の終了時にロギング後処理をしたいときに使用
    # 2. 呼び元: with構文の終了時に自動で呼ばれることを想定
    # 3. 内容: ログファイルのパス出力など後処理を行う
    def __exit__(self, exc_type, exc_value, traceback):
        self.finalize_logging()
        return False

    # 1. シーン: コンソールハンドラに色付きフォーマッタを適用したいときに使用
    # 2. 呼び元: initialize_logging()からのみ呼ばれることを想定
    # 3. 内容: colorlogを使ってログレベルごとに色分けしたフォーマッタを設定する
    def _apply_colorlog_to_console_handler(self, handler):
        """
        コンソールハンドラーに colorlog を使ったフォーマッタを適用する。
        """
        log_colors = {
            "DEBUG": "green",  # 緑
            "INFO": "black",  # 黒
            "WARNING": "red",  # 赤
            "ERROR": "white,bg_red",  # 白、背景赤
            "CRITICAL": "yellow,bg_red",  # 黄色、背景赤
        }
        formatter = colorlog.ColoredFormatter(
            "%(log_color)s[%(thread)s] %(levelname)s: %(message)s%(reset)s",  # フォーマットを修正
            log_colors=log_colors,
        )
        handler.setFormatter(formatter)

    # 1. シーン: ロギングの初期化を行いたいときに使用
    # 2. 呼び元: __enter__()やmain関数などから呼ばれることを想定
    # 3. 内容: ファイル・コンソール両方のハンドラをセットアップし色付けも行う
    def initialize_logging(self):
        self.base_time = datetime.now()

        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)  # これは変更しない！

        logform_stdout = "[%(thread)s] %(levelname)s: %(message)s"
        streamH = logging.StreamHandler()
        streamH.setFormatter(logging.Formatter(logform_stdout))
        streamH.setLevel(logging.INFO)  # これは変更しない！
        logger.addHandler(streamH)
        self._stream_handler = streamH

        # ここで色付きフォーマッタを適用する
        self._apply_colorlog_to_console_handler(streamH)

        self.set_stream_level(self._level)  # ここで調節！

    # 1. シーン: コンソール出力のログレベルを動的に変更したいときに使用
    # 2. 呼び元: main関数やSmartLogging利用箇所から呼ばれることを想定
    # 3. 内容: ストリームハンドラのログレベルを指定値に変更する
    def set_stream_level(self, level: int) -> None:
        if not self._stream_handler:
            return
        self._stream_handler.setLevel(level)

    # 1. シーン: ロギングの終了時にファイルパスなどを出力したいときに使用
    # 2. 呼び元: __exit__()やmain関数などから呼ばれることを想定
    # 3. 内容: ファイルハンドラのログファイルパスを相対パスで出力する
    def finalize_logging(self):
        for handler in logging.getLogger().handlers:
            if isinstance(handler, logging.FileHandler):
                rel = Path(os.path.relpath(handler.baseFilename, start=Path.cwd()))
                elapsed = int((datetime.now() - self.base_time).total_seconds())
                m, s = divmod(elapsed, 60)
                h, m = divmod(m, 60)
                logging.info(
                    f"Time: {h}h {m:02d}min {s:02d}sec"
                    if h
                    else f"Time: {m}min {s:02d}sec" if m else f"Time: {s}sec"
                )
                logging.info(f"Log: {rel.as_posix()}")

    # 1. シーン: プロジェクト配下のログだけを通すフィルタを定義したいときに利用する内部クラス
    # 2. 呼び元: set_stream_filter()からのみ利用されることを想定
    # 3. 内容: ログレコードのパスがプロジェクト配下かどうかでフィルタリングする
    class _OnlyMyLogsFilter(logging.Filter):
        """プロジェクト配下のログだけを通す内部用フィルター"""

        def filter(self, record: logging.LogRecord) -> bool:
            try:
                pathname = Path(record.pathname).resolve()
                return PROJ_ABSPATH in pathname.parents or pathname == PROJ_ABSPATH
            except Exception:
                return False

    # 1. シーン: コンソール出力のログにプロジェクト配下のみのフィルタを付与したいときに使用
    # 2. 呼び元: main関数やSmartLogging利用箇所から呼ばれることを想定
    # 3. 内容: ストリームハンドラにOnlyMyLogsFilterを付与・除去する
    def set_stream_filter(self, filter_on: bool) -> None:
        """STDOUTのログにOnlyMyLogsFilterをつけたり外したりする"""
        if not self._stream_handler:
            return

        self._stream_handler.filters = [
            f
            for f in self._stream_handler.filters
            if not isinstance(f, SmartLogging._OnlyMyLogsFilter)
        ]
        if filter_on:
            self._stream_handler.addFilter(SmartLogging._OnlyMyLogsFilter())
