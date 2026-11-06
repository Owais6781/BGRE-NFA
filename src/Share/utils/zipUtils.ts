import JSZip from 'jszip';

export interface IZipFileItem {
  fileName: string;
  serverRelativeUrl?: string;
}

/**
 * Downloads multiple files from SharePoint and packages them into a single ZIP file.
 * @param files List of files with fileName and serverRelativeUrl
 * @param zipFileName Name of the downloaded zip file (e.g. 'NFA_Attachments.zip')
 * @param onProgress Optional progress callback (completed, total)
 */
export async function downloadFilesAsZip(
  files: IZipFileItem[],
  zipFileName: string = 'Attachments.zip',
  onProgress?: (completed: number, total: number) => void
): Promise<void> {
  const validFiles = (files || []).filter(f => f && f.fileName && f.serverRelativeUrl);
  if (validFiles.length === 0) {
    alert('No files available to download.');
    return;
  }

  const zip = new JSZip();
  let completed = 0;
  const fileNameCounts: { [name: string]: number } = {};

  for (const file of validFiles) {
    try {
      // Avoid filename collision inside the ZIP
      let uniqueName = file.fileName;
      if (fileNameCounts[uniqueName]) {
        const dotIdx = uniqueName.lastIndexOf('.');
        const namePart = dotIdx !== -1 ? uniqueName.substring(0, dotIdx) : uniqueName;
        const extPart = dotIdx !== -1 ? uniqueName.substring(dotIdx) : '';
        uniqueName = `${namePart} (${fileNameCounts[file.fileName]})${extPart}`;
        fileNameCounts[file.fileName]++;
      } else {
        fileNameCounts[uniqueName] = 1;
      }

      const fileUrl = file.serverRelativeUrl!;
      const fullUrl = fileUrl.startsWith('http')
        ? fileUrl
        : `${window.location.origin}${fileUrl.startsWith('/') ? '' : '/'}${fileUrl}`;

      const res = await fetch(fullUrl, {
        method: 'GET',
        credentials: 'same-origin',
        headers: {
          Accept: '*/*'
        }
      });

      if (!res.ok) {
        console.warn(`Failed to fetch file: ${file.fileName} (${res.status} ${res.statusText})`);
        continue;
      }

      const blob = await res.blob();
      zip.file(uniqueName, blob);
    } catch (err) {
      console.warn(`Error fetching file for ZIP: ${file.fileName}`, err);
    } finally {
      completed++;
      if (onProgress) {
        onProgress(completed, validFiles.length);
      }
    }
  }

  const zipFiles = Object.keys(zip.files);
  if (zipFiles.length === 0) {
    alert('No files could be downloaded to create the ZIP.');
    return;
  }

  const zipBlob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });

  const downloadUrl = window.URL.createObjectURL(zipBlob);
  const anchor = document.createElement('a');
  anchor.href = downloadUrl;
  const cleanZipName = zipFileName.toLowerCase().endsWith('.zip') ? zipFileName : `${zipFileName}.zip`;
  anchor.download = cleanZipName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(() => window.URL.revokeObjectURL(downloadUrl), 5000);
}
