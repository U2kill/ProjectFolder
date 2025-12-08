import cv2
import os
from pathlib import Path
from PySide6.QtCore import (
    QObject,
    QRunnable,
    QThreadPool,
    QTimer,
    Signal,
    Slot,
)

class WorkerSignals(QObject):
    """Signals from a running worker thread.

    finished
        No data

    error
        tuple (exctype, value, traceback.format_exc())

    result
        object data returned from processing, anything

    progress
        float indicating % progress
    """

    finished = Signal()
    error = Signal(tuple)
    result = Signal(str)
    progress = Signal(str)


class Converter(QRunnable):
    def __init__(self, movPaths, savePath):
        super().__init__()
        self.movPaths = movPaths
        self.savePath = savePath
        self.signals = WorkerSignals()


    def convert_mov_to_mp4_opencv_basic(self, input_path, output_path):
        """
        Базовая конвертация MOV в MP4 с помощью OpenCV
        ВНИМАНИЕ: Аудио не сохраняется!
        """
        # Открываем видеофайл
        cap = cv2.VideoCapture(input_path)
        
        if not cap.isOpened():
            print(f"Ошибка: не удалось открыть файл {input_path}")
            return False
        
        # Получаем параметры исходного видео
        frame_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        frame_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        fps = cap.get(cv2.CAP_PROP_FPS)
        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        
        self.signals.progress.emit(f"\nИсходное видео:")
        self.signals.progress.emit(f"  Название: {Path(input_path).stem}.MOV")
        self.signals.progress.emit(f"  Разрешение: {frame_width}x{frame_height}")
        self.signals.progress.emit(f"  FPS: {fps:.2f}")
        self.signals.progress.emit(f"  Длительность: {total_frames/fps:.2f} сек")

        
        # Определяем кодек для MP4
        # Четырехсимвольный код кодека:
        # 'mp4v' - MPEG-4 кодек

        fourcc = cv2.VideoWriter_fourcc(*'mp4v')
        
        output_file_path = Path(output_path) / f"{Path(input_path).stem}.mp4"

        # Создаем объект для записи видео
        out = cv2.VideoWriter(
            output_file_path,
            fourcc,
            fps,
            (frame_width, frame_height)
        )
        
        frame_count = 0
        
        self.signals.progress.emit("Начало конвертации...")
        
        while True:
            ret, frame = cap.read()
            
            if not ret:
                break
            
            # Записываем кадр
            out.write(frame)
            frame_count += 1
            
            # Выводим прогресс каждые 100 кадров
            if frame_count % 100 == 0:
                progress = (frame_count / total_frames) * 100
                self.signals.progress.emit(f"Прогресс: {progress:.1f}% ({frame_count}/{total_frames})")
        
        # Освобождаем ресурсы
        cap.release()
        out.release()
        cv2.destroyAllWindows()
        
        self.signals.progress.emit(f"Конвертация завершена!")
        # self.signals.progress.emit(f"Обработано кадров: {frame_count}")
        
        # Проверяем размер файла
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path) / (1024 * 1024)
            self.signals.progress.emit(f"Размер выходного файла: {file_size:.2f} MB")
        
        return True

    @Slot()
    def run(self):
        for i in self.movPaths:
            self.convert_mov_to_mp4_opencv_basic(i, self.savePath)
        
        self.signals.finished.emit()
