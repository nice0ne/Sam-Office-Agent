import { ChunkingOptions, DocumentChunk } from './types';

export const ID_STOP_WORDS = new Set([
  'ada', 'adalah', 'adanya', 'adapun', 'agak', 'agar', 'akan', 'akankah', 'akhir',
  'akhiri', 'akhirnya', 'aku', 'akulah', 'amat', 'amatlah', 'anda', 'andalah',
  'antar', 'antara', 'antaranya', 'apa', 'apaan', 'apabila', 'apakah', 'apalagi',
  'apatah', 'artinya', 'asal', 'asalkan', 'atas', 'atau', 'ataukah', 'ataupun',
  'awal', 'awalnya', 'bagai', 'bagaikan', 'bagaimana', 'bagaimanakah', 'bagaimanapun',
  'bagi', 'bahkan', 'bahwa', 'bahwasanya', 'baik', 'banyak', 'banyaknya', 'beberapa',
  'begini', 'beginian', 'beginikah', 'beginilah', 'begitu', 'begitukah', 'begitulah',
  'begitupun', 'belum', 'belumlah', 'berada', 'berapa', 'berapakah', 'berapalah',
  'berapapun', 'berarti', 'berikut', 'berikutnya', 'berjumlah', 'berkenaan', 'berlainan',
  'bersama', 'berujar', 'besar', 'betul', 'biasa', 'biasanya', 'bila', 'bilakah',
  'bisa', 'bisakah', 'boleh', 'bolehlah', 'buat', 'bukan', 'bukankah', 'bukanlah',
  'cuma', 'dahulu', 'dalam', 'dan', 'dapat', 'dari', 'daripada', 'dekat', 'demi',
  'demikian', 'demikianlah', 'dengan', 'depan', 'di', 'dia', 'dialah', 'diantara',
  'diantaranya', 'dikarenakan', 'dini', 'diri', 'dirinya', 'disini', 'disinilah',
  'dituturkan', 'dua', 'dulu', 'empat', 'enggak', 'hal', 'hanya', 'harus', 'haruslah',
  'hampir', 'hari', 'harinya', 'hingga', 'ia', 'ialah', 'ibarat', 'ingin', 'inginkah',
  'ini', 'inikah', 'inilah', 'itu', 'itukah', 'itulah', 'jadi', 'jangan', 'jika',
  'jikalau', 'juga', 'jumlah', 'justru', 'kala', 'kalau', 'kalian', 'kami', 'kamilah',
  'kamu', 'kamulah', 'kan', 'kapan', 'kapanpun', 'karena', 'karenanya', 'kata',
  'katanya', 'ke', 'keadaan', 'kecil', 'kembali', 'kemudian', 'kenapa', 'kepada',
  'kepadanya', 'keseluruhan', 'khususnya', 'kini', 'kinilah', 'kira', 'kiranya',
  'kita', 'kitalah', 'kurang', 'lagi', 'lagian', 'lah', 'lain', 'lainnya', 'lalu',
  'lama', 'lamanya', 'lebih', 'lewat', 'macam', 'maka', 'makanya', 'makin', 'malah',
  'malahan', 'mampu', 'mana', 'manakala', 'manalagi', 'masa', 'masalah', 'masih',
  'masihkah', 'masing', 'mau', 'maupun', 'melainkan', 'melalui', 'memang', 'mengapa',
  'mengenai', 'menjadi', 'menurut', 'mereka', 'merekalah', 'merupakan', 'meski',
  'meskipun', 'mungkin', 'nah', 'namun', 'nanti', 'nyata', 'oleh', 'olehnya', 'pada',
  'padahal', 'padanya', 'paling', 'panjang', 'pantas', 'para', 'pasti', 'pastilah',
  'penting', 'per', 'pernah', 'persoalan', 'pertama', 'pihak', 'pula', 'pun', 'punya',
  'rasa', 'saat', 'saatnya', 'saja', 'sajalah', 'saling', 'sama', 'sambil', 'sampai',
  'sana', 'sangat', 'sangatlah', 'satu', 'saya', 'sayalah', 'sebab', 'sebagai',
  'sebagaimana', 'sebagian', 'sebaik', 'sebaliknya', 'sebanyak', 'sebegini', 'sebegitu',
  'sebelum', 'sebelumnya', 'sebenarnya', 'seberapa', 'sebesar', 'sebetulnya', 'secara',
  'sedang', 'sedangkan', 'sedikit', 'sehingga', 'sejak', 'sejauh', 'sekali', 'sekalian',
  'sekaligus', 'sekalipun', 'sekarang', 'sekitar', 'selain', 'selaku', 'selalu',
  'selama', 'selamanya', 'seluruh', 'semacam', 'semakin', 'sementara', 'semisal',
  'semua', 'semula', 'sendiri', 'sendirinya', 'seolah', 'seorang', 'sepanjang',
  'seperti', 'sering', 'serta', 'serupa', 'sesaat', 'sesama', 'sesuatu', 'sesudah',
  'sesudahnya', 'setelah', 'setempat', 'setengah', 'seterusnya', 'setiap', 'setiba',
  'seusai', 'sewaktu', 'siap', 'siapa', 'siapakah', 'siapapun', 'sini', 'sinilah',
  'suatu', 'sudah', 'sudahkah', 'sudahlah', 'supaya', 'tadi', 'tadinya', 'tahu',
  'tahun', 'tak', 'tanpa', 'tapi', 'telah', 'tentang', 'tentu', 'tentulah', 'tentunya',
  'tepat', 'terhadap', 'terhadapnya', 'termasuk', 'tersebut', 'tertentu', 'tetapi',
  'tiap', 'tidak', 'tidakkah', 'tidaklah', 'toh', 'waduh', 'wah', 'wahai', 'waktu',
  'walau', 'walaupun', 'wong', 'ya', 'yaitu', 'yakin', 'yakni', 'yang'
]);

export const EN_STOP_WORDS = new Set([
  'a', 'about', 'above', 'after', 'again', 'against', 'all', 'am', 'an', 'and',
  'any', 'are', 'aren', 'as', 'at', 'be', 'because', 'been', 'before', 'being',
  'below', 'between', 'both', 'but', 'by', 'can', 'cannot', 'could', 'did', 'do',
  'does', 'doing', 'down', 'during', 'each', 'few', 'for', 'from', 'further',
  'had', 'has', 'have', 'having', 'he', 'her', 'here', 'hers', 'herself', 'him',
  'himself', 'his', 'how', 'i', 'if', 'in', 'into', 'is', 'it', 'its', 'itself',
  'let', 'me', 'more', 'most', 'my', 'myself', 'no', 'nor', 'not', 'of', 'off',
  'on', 'once', 'only', 'or', 'other', 'ought', 'our', 'ours', 'ourselves', 'out',
  'over', 'own', 'same', 'she', 'should', 'so', 'some', 'such', 'than', 'that',
  'the', 'their', 'theirs', 'them', 'themselves', 'then', 'there', 'these',
  'they', 'this', 'those', 'through', 'to', 'too', 'under', 'until', 'up', 'very',
  'was', 'we', 'were', 'what', 'when', 'where', 'which', 'while', 'who', 'whom',
  'why', 'with', 'would', 'you', 'your', 'yours', 'yourself', 'yourselves'
]);

export function extractFileType(fileName: string): string {
  if (!fileName) return 'txt';
  const parts = fileName.split('.');
  if (parts.length > 1) {
    return parts[parts.length - 1].toLowerCase();
  }
  return 'txt';
}

export function extractKeywords(text: string): string[] {
  if (!text) return [];
  const words = text.toLowerCase().match(/[a-z0-9_\u00C0-\u017F]+/g) || [];
  const seen = new Set<string>();
  const keywords: string[] = [];

  for (const word of words) {
    if (word.length <= 1 || /^\d+$/.test(word)) {
      continue;
    }
    if (ID_STOP_WORDS.has(word) || EN_STOP_WORDS.has(word)) {
      continue;
    }
    if (!seen.has(word)) {
      seen.add(word);
      keywords.push(word);
    }
  }

  return keywords;
}

function getOverlapText(text: string, overlapSize: number): string {
  if (overlapSize <= 0 || !text) return '';
  if (text.length <= overlapSize) return text.trim();
  const rawOverlap = text.slice(-overlapSize);
  const spaceIdx = rawOverlap.indexOf(' ');
  if (spaceIdx !== -1 && spaceIdx < rawOverlap.length - 1) {
    return rawOverlap.slice(spaceIdx + 1).trim();
  }
  return rawOverlap.trim();
}

function splitLongText(text: string, targetSize: number): string[] {
  const words = text.split(/\s+/).filter(Boolean);
  const result: string[] = [];
  let current = '';

  for (const word of words) {
    if (word.length > targetSize) {
      if (current) {
        result.push(current);
        current = '';
      }
      for (let i = 0; i < word.length; i += targetSize) {
        result.push(word.slice(i, i + targetSize));
      }
    } else if (!current) {
      current = word;
    } else if ((current + ' ' + word).length <= targetSize) {
      current += ' ' + word;
    } else {
      result.push(current);
      current = word;
    }
  }

  if (current) {
    result.push(current);
  }

  return result.length > 0 ? result : [text];
}

function splitIntoParagraphs(text: string, targetSize: number): string[] {
  const rawParagraphs = text.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);
  const segments: string[] = [];

  for (const para of rawParagraphs) {
    if (para.length <= targetSize) {
      segments.push(para);
    } else {
      const sentences = para.match(/[^.!?\n]+[.!?\n]+|[^.!?\n]+$/g) || [para];
      let currentSentenceChunk = '';
      for (const sentence of sentences) {
        const trimmed = sentence.trim();
        if (!trimmed) continue;
        if (trimmed.length > targetSize) {
          if (currentSentenceChunk) {
            segments.push(currentSentenceChunk);
            currentSentenceChunk = '';
          }
          segments.push(...splitLongText(trimmed, targetSize));
        } else if (!currentSentenceChunk) {
          currentSentenceChunk = trimmed;
        } else if ((currentSentenceChunk + ' ' + trimmed).length <= targetSize) {
          currentSentenceChunk += ' ' + trimmed;
        } else {
          segments.push(currentSentenceChunk);
          currentSentenceChunk = trimmed;
        }
      }
      if (currentSentenceChunk) {
        segments.push(currentSentenceChunk);
      }
    }
  }

  return segments.length > 0 ? segments : [text.trim()];
}

export function chunkText(
  text: string,
  fileName: string,
  options?: ChunkingOptions,
  docId?: string
): DocumentChunk[] {
  if (!text || !text.trim()) {
    return [];
  }

  const targetChunkSize = options?.targetChunkSize ?? 1200;
  const chunkOverlap = options?.chunkOverlap ?? 150;
  const fileType = extractFileType(fileName);
  const resolvedDocId = docId || `doc_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;

  const segments = splitIntoParagraphs(text, targetChunkSize);
  const rawChunks: string[] = [];
  let currentChunk = '';

  for (const segment of segments) {
    if (!currentChunk) {
      currentChunk = segment;
    } else if ((currentChunk + '\n\n' + segment).length <= targetChunkSize) {
      currentChunk += '\n\n' + segment;
    } else {
      rawChunks.push(currentChunk);
      const overlapText = getOverlapText(currentChunk, chunkOverlap);
      currentChunk = overlapText ? `${overlapText}\n\n${segment}` : segment;
    }
  }

  if (currentChunk.trim()) {
    rawChunks.push(currentChunk.trim());
  }

  return rawChunks.map((chunkStr, index) => {
    const trimmed = chunkStr.trim();
    return {
      id: `${resolvedDocId}_chunk_${index}`,
      documentId: resolvedDocId,
      fileName,
      fileType,
      chunkIndex: index,
      text: trimmed,
      charCount: trimmed.length,
      wordCount: trimmed.split(/\s+/).filter(Boolean).length,
      keywords: extractKeywords(trimmed),
    };
  });
}
