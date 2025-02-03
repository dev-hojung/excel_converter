import { useState } from "react";

export const useConvertAggregation = () => {
 
  const [selectOption, setSelectOption] = useState<'yearly' | 'monthly'>('yearly')
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);

  return {
    file,
    downloadUrl,
    isLoading,
    selectOption,
    setFile,
    setDownloadUrl,
    setIsLoading,
    setSelectOption
  }
}