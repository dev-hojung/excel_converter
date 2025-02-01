import React, { useState, DragEvent } from 'react';

export default function UploadPage() {
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>('');
  const [isLoading, setIsLoading] = useState(false); // 로딩 상태

  // 파일 직접 선택 시
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
    e.target.value = '';
  };

  // 드래그된 파일을 영역에 놓을 때
  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  // 드래그가 영역 위를 지날 때 (기본 이벤트 막아야 drop 가능)
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  // 파일 제거
  const handleRemoveFile = () => {
    setFile(null);
    setDownloadUrl('');
  };

  // 업로드 & 변환
  const handleUpload = async () => {
    if (!file) {
      alert('파일을 먼저 선택(또는 드래그)하세요.');
      return;
    }

    try {
      setIsLoading(true); // 업로드 시작 시 로딩 표시
      const formData = new FormData();
      formData.append('excel', file);

      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error(`서버 응답 에러: ${response.statusText}`);
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      alert('변환이 완료되었습니다. 다운로드 버튼을 클릭하세요!');
    } catch (error) {
      console.error(error);
      alert('업로드/변환 중 오류가 발생했습니다.');
    } finally {
      setIsLoading(false); // 업로드 끝
    }
  };

  // 다운로드
  const handleDownload = () => {
    if (!downloadUrl) return;
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = 'custom.xlsx'; // 다운로드 파일명
    link.click();
  };

  return (
    <div className="relative max-w-xl mx-auto py-8">
      {/* 전체 화면 로딩 오버레이 */}
      {isLoading && (
  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
    <div
      className="flex flex-col items-center p-8 rounded-md shadow-xl"
      style={{
        background: 'linear-gradient(135deg, #ffffff 0%, #f0f4f8 100%)',
      }}
    >
      <span className="inline-block h-16 w-16 animate-spin rounded-full border-8 border-purple-500 border-t-transparent mb-6"></span>
      <p className="text-lg text-gray-800 font-semibold">처리 중...</p>
    </div>
  </div>
)}

      <h1 className="text-2xl font-bold mb-6 text-center">엑셀 업로드 및 변환</h1>

      {/* 드래그앤드롭 영역 */}
      <div
        className="w-full p-6 mb-4 text-center border-2 border-dashed border-gray-300 rounded-lg hover:border-blue-300 transition-colors cursor-pointer"
        onDragOver={handleDragOver}
        onDrop={handleDrop}
      >
        {file ? (
          <p className="text-gray-700 font-medium">
            {file.name} (선택됨)
          </p>
        ) : (
          <p className="text-gray-500">
            이 영역에 파일을 드래그&드롭 하거나,<br />
            아래 버튼으로 파일을 선택하세요.
          </p>
        )}
      </div>

      {/* 파일 선택 & 제거 버튼들 */}
      <div className="flex items-center gap-2 mb-4">
        <label
          htmlFor="excel-file"
          className="inline-block px-4 py-2 text-white bg-blue-500 rounded-md cursor-pointer hover:bg-blue-600 transition-colors"
        >
          파일 선택
        </label>
        <input
          id="excel-file"
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileChange}
          className="hidden"
        />

        {file && (
          <button
            onClick={handleRemoveFile}
            className="px-4 py-2 bg-red-500 text-white rounded-md hover:bg-red-600 transition-colors"
          >
            파일 제거
          </button>
        )}
      </div>

      {/* 업로드 & 변환 버튼 */}
      <button
        onClick={handleUpload}
        className="px-4 py-2 mr-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
      >
        업로드 및 변환
      </button>

      {/* 다운로드 버튼 */}
      {downloadUrl && (
        <button
          onClick={handleDownload}
          className="px-4 py-2 bg-indigo-500 text-white rounded-md hover:bg-indigo-600 transition-colors"
        >
          결과 다운로드
        </button>
      )}
    </div>
  );
}