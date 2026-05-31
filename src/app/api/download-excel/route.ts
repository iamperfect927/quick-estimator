import { NextResponse } from 'next/server';
import { generateEstimateExcel } from '@/app/libs/excelGenerator';

export async function POST(req: Request) {
  try {
    const data = await req.json();
    const buffer = await generateEstimateExcel(data);
    const fileBytes = new Uint8Array(buffer);

    return new NextResponse(fileBytes, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="estimate.xlsx"',
      },
    });
  } catch (error) {
    return NextResponse.json({ error: 'Failed to generate sheet' }, { status: 500 });
  }
}