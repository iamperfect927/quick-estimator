import { NextRequest, NextResponse } from 'next/server';

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);
  const fileUrl = searchParams.get('url');

  if (!fileUrl) {
    return NextResponse.json({ error: 'Missing "url" parameter' }, { status: 400 });
  }

  // Regular expression to match standard Google Drive share links and export links
  const dMatch = fileUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  const idMatch = fileUrl.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  const fileId = dMatch?.[1] || idMatch?.[1];

  if (!fileId) {
    return NextResponse.json({ 
      error: 'Invalid Google Drive link pattern. Make sure it contains a valid file ID.' 
    }, { status: 400 });
  }

  // 1. Google Sheets export URL (translates Sheets directly to an Excel file stream)
  const sheetsExportUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
  
  // 2. Direct Google Drive file download link (works for uploaded standard Excel files)
  const driveDownloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

  try {
    console.log(`[Google Drive Proxy] Attempting Google Sheets XLSX export for ID: ${fileId}`);
    let response = await fetch(sheetsExportUrl);

    // If it's not a Google Sheet, or the export failed, try direct file download
    const contentType = response.headers.get('content-type') || '';
    if (!response.ok || contentType.includes('text/html')) {
      console.log(`[Google Drive Proxy] Sheets export failed or returned HTML. Trying direct drive download for ID: ${fileId}`);
      response = await fetch(driveDownloadUrl);
    }

    if (!response.ok) {
      return NextResponse.json({ 
        error: `Could not retrieve file. Verify that the file is shared with "Anyone with the link can view".` 
      }, { status: 400 });
    }

    const buffer = await response.arrayBuffer();

    // Determine filename
    const contentDisp = response.headers.get('content-disposition') || '';
    let filename = `drive_file_${fileId}.xlsx`;
    const filenameMatch = contentDisp.match(/filename="?([^"]+)"?/);
    if (filenameMatch && filenameMatch[1]) {
      filename = filenameMatch[1];
    } else if (response.url.includes('docs.google.com/spreadsheets')) {
      filename = `Google_Sheet_${fileId}.xlsx`;
    }

    return new NextResponse(new Uint8Array(buffer), {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${filename}"`,
      }
    });

  } catch (error: any) {
    console.error('[Google Drive Proxy] Error fetching from Google Drive:', error);
    return NextResponse.json({ error: `Connection failed: ${error.message}` }, { status: 500 });
  }
}
