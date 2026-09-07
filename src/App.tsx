import React, { useState, useEffect } from 'react';
import { GoogleGenAI } from '@google/genai';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import JSZip from 'jszip';
import { 
  GraduationCap, 
  Key, 
  Settings, 
  BookOpen, 
  CheckCircle2, 
  UploadCloud, 
  SlidersHorizontal, 
  Heart, 
  Globe, 
  Languages, 
  ExternalLink,
  Sparkles,
  Check,
  Pin,
  Download,
  Copy,
  ChevronUp,
  ChevronDown,
  Calculator,
  AlertCircle,
  Zap,
  Save,
  X
} from 'lucide-react';

export default function App() {
  const [showResult, setShowResult] = useState(false);
  const [lessonPlanFile, setLessonPlanFile] = useState<File | null>(null);
  const [curriculumFile, setCurriculumFile] = useState<File | null>(null);
  const [textbookFile, setTextbookFile] = useState<File | null>(null);

  // New states for form, API key, and processing
  const [showApiModal, setShowApiModal] = useState(false);
  const [apiKey, setApiKey] = useState(() => localStorage.getItem('user_gemini_api_key') || '');
  const [aiModel, setAiModel] = useState(() => localStorage.getItem('user_gemini_model') || 'gemini-3.6-flash');
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [aiResultText, setAiResultText] = useState("");
  const [subject, setSubject] = useState('Toán');
  const [grade, setGrade] = useState('Lớp 12');
  const [options, setOptions] = useState({
    ai: false,
    inclusion: false,
    language: false,
    bilingual: false,
    toanKTD: false
  });
  const [aiResult, setAiResult] = useState<{mucTieu: string, hoatDong: string} | null>(null);
  const [errorMsg, setErrorMsg] = useState("");
  const [progressText, setProgressText] = useState("Đang phân tích cấu trúc giáo án...");

  const getProgressText = (p: number) => {
    if (p < 25) return "Đang phân tích cấu trúc giáo án...";
    if (p < 40) return "Đang xác định các hoạt động dạy học...";
    if (p < 55) return "Đang lập bản đồ nội dung giáo án...";
    if (p < 70) return "Đang lựa chọn vị trí chèn NLS...";
    if (p < 85) return "Đang tạo nội dung NLS chi tiết...";
    return "Đang hoàn thiện giáo án...";
  };

  const handleStart = async () => {
    setErrorMsg("");
    if (!lessonPlanFile) {
      alert("Vui lòng tải lên file giáo án (bắt buộc)!");
      return;
    }
    
    if (!lessonPlanFile.name.endsWith('.docx')) {
      setErrorMsg("Hệ thống chỉ hỗ trợ xử lý file Word định dạng .docx. Vui lòng chuyển đổi file của bạn sang .docx và tải lên lại.");
      return;
    }
    
    if (!apiKey) {
      setShowApiModal(true);
      return;
    }

    setIsProcessing(true);
    setShowResult(false);
    setProgress(0);
    setProgressText("Đang phân tích cấu trúc giáo án...");
    setAiResultText("");
    
    const progressInterval = setInterval(() => {
      setProgress(p => {
        const next = p + Math.floor(Math.random() * 5) + 2;
        const newProgress = next >= 95 ? 95 : next;
        setProgressText(getProgressText(newProgress));
        return newProgress;
      });
    }, 800);
    
    try {
      const ai = new GoogleGenAI({ apiKey: apiKey.trim() });
      
      const extractTextFromDocx = async (file: File) => {
        try {
          const zip = new JSZip();
          const loadedZip = await zip.loadAsync(file);
          const docXml = await loadedZip.file("word/document.xml")?.async("string");
          if (docXml) {
            return docXml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
          }
        } catch (e) {
          console.error("Lỗi khi đọc file docx:", e);
        }
        return "Không thể đọc nội dung file.";
      };

      // Extract text from the main lesson plan docx
      let fileTextContent = await extractTextFromDocx(lessonPlanFile);
      fileTextContent = fileTextContent.substring(0, 30000); // Limit length

      let prompt = `Đóng vai chuyên gia công nghệ giáo dục và sư phạm. Tôi đang soạn giáo án môn ${subject} ${grade}.
      Bạn hãy phân tích file giáo án đầu vào và trả về văn bản chứa CÁC ĐOẠN TEXT CẦN CHÈN để tôi bổ sung vào file Word.
      
      YÊU CẦU QUAN TRỌNG VỀ CẤU TRÚC VÀ VỊ TRÍ CHÈN (TUYỆT ĐỐI TUÂN THỦ CÁC QUY TẮC SAU):
      Quy tắc 1: CHÈN MỤC TIÊU CHUNG (2.3, 2.4, 2.5) CHÍNH XÁC VÀO CUỐI PHẦN "2. VỀ NĂNG LỰC".
      - TUYỆT ĐỐI KHÔNG được chèn các mục này ở cuối bài! Bạn BẮT BUỘC phải tìm đúng phần "I. MỤC TIÊU" -> "2. Về năng lực" (hoặc mục tương đương ở đầu giáo án).
      - Hãy tìm CÂU VĂN HOẶC ĐOẠN VĂN CUỐI CÙNG của phần "2. Về năng lực" (đoạn nằm ngay trước khi bắt đầu sang phần "3. Về phẩm chất" hoặc "II. THIẾT BỊ DẠY HỌC").
      - Copy Y HỆT NGUYÊN VĂN câu văn/đoạn văn cuối cùng đó để làm TARGET. Hệ thống sẽ chèn NLS ngay bên dưới câu văn này.
      - Chú ý: KHÔNG dùng tiêu đề "2. Về năng lực" làm TARGET vì nó sẽ chèn lộn lên trên cùng của phần này. Hãy dùng nội dung của ý cuối cùng trong phần 2 làm TARGET.
      - Ngay sau TARGET này, chèn các mục: "2.3. Năng lực số", "2.4. Lồng ghép AI" (nếu có), "2.5. Lồng ghép STEM" (nếu có). Bạn PHẢI COPY NGUYÊN VĂN toàn bộ nội dung năng lực số từ file PPCT.
      - TUYỆT ĐỐI KHÔNG chèn mục 2.3, 2.4, 2.5 vào các Hoạt động (Hoạt động 1, Hoạt động 2...) ở bên dưới. Nó phải nằm ở phần MỤC TIÊU CHUNG của toàn bài.
      
      Quy tắc 2: CÁCH VIẾT MỤC TIÊU VÀ SỰ NHẤT QUÁN. 
      - Khi chèn mục tiêu tích hợp vào một Hoạt động, BẠN PHẢI GHI RÕ THEO ĐÚNG CẤU TRÚC SAU: "- Tích hợp [Tên nội dung tích hợp]: [Mã năng lực nếu có từ PPCT]: [Hành động] để [Mục đích/Kết quả đạt được]."
      - Ví dụ chuẩn: "- Tích hợp năng lực số: 1.1.NC1b: Biết ứng dụng máy tính cầm tay (chế độ Thống kê) để tính nhanh phương sai $s^2$ và độ lệch chuẩn $s$ của mẫu số liệu ghép nhóm."
      - QUAN TRỌNG: Nếu trong PPCT hoặc giáo án KHÔNG CÓ mục đích (tức là không có phần "để làm gì"), AI PHẢI TỰ SUY LUẬN dựa vào nội dung bài học để tự điền thêm phần "để..." vào cuối câu. Nếu đã có sẵn thì giữ nguyên. Tuyệt đối không chỉ ghi mã NLS cộc lốc hoặc thiếu mục đích.
      - NẾU bạn thêm mục tiêu Năng lực số/AI vào một Hoạt động (ví dụ HĐ 1), thì BẮT BUỘC trong phần "Tổ chức thực hiện" của CHÍNH HOẠT ĐỘNG ĐÓ phải có hành động tương ứng của GV/HS. Tuyệt đối không được chèn mục tiêu ở HĐ 1 nhưng lại nhét hành động vào bảng của HĐ 2.
      
      Quy tắc 3: CÁCH XÁC ĐỊNH [TARGET] CHÍNH XÁC 100%:
      - ĐỂ CHÈN VÀO MỤC TIÊU CHUNG ("2. Về năng lực"): TUYỆT ĐỐI KHÔNG dùng các tiêu đề (như "2. Về năng lực" hoặc "Mục tiêu") làm TARGET, vì nó sẽ chèn nội dung lên trên cùng, đẩy các nội dung cũ xuống dưới. Bạn PHẢI COPY NGUYÊN VĂN CÂU VĂN CUỐI CÙNG của mục đó để làm TARGET. Hệ thống sẽ chèn vào ngay bên dưới câu đó.
      - ĐỂ CHÈN VÀO MỤC TIÊU CỦA TỪNG HOẠT ĐỘNG (Ví dụ: "a) Mục tiêu" trong Hoạt động 1): Tương tự, hãy copy câu văn cuối cùng của phần "a) Mục tiêu" đó làm TARGET.
      - ĐỂ CHÈN VÀO BẢNG "TỔ CHỨC THỰC HIỆN": TUYỆT ĐỐI KHÔNG dùng các từ ở cột trái (như "Chuyển giao", "Thực hiện", "Báo cáo") làm TARGET, vì nội dung sẽ bị chèn nhầm vào cột trái. 
      - BẠN PHẢI copy TRỌN VẸN MỘT CÂU VĂN ở CỘT PHẢI (cột nội dung) của bước tương ứng làm TARGET. Hệ thống sẽ chèn NLS vào ngay bên dưới câu văn đó ở cột phải. (Ví dụ: [TARGET: GV yêu cầu HS thảo luận nhóm đôi để trả lời các câu hỏi.]).
      Vì các khối TARGET được xử lý theo thứ tự từ trên xuống dưới, nên hệ thống sẽ tự động tìm đúng vị trí của từng hoạt động.
      
      Quy tắc 4: Mọi văn bản bạn sinh ra BẮT BUỘC PHẢI BỌC TRONG THẺ MÀU. Dùng [RED]...[/RED] cho Năng lực số/STEM và [BLUE]...[/BLUE] cho Năng lực AI. Cả những mục như "2.3. Năng lực số:" cũng phải bọc thẻ. Nếu không bọc thẻ, chữ sẽ bị màu đen.
      
      Quy tắc 5: Nếu có công thức toán học, HÃY GÕ DƯỚI DẠNG CÔNG THỨC LATEX (ví dụ: $x^2$, $\frac{a}{b}$, $\sigma$) theo đúng yêu cầu của người dùng.
      
      Quy tắc 6 (Vận dụng STEM): Nếu tùy chọn "Toán KTD (STEM)" được bật, hãy kiểm tra xem trong file PPCT có đề cập đến chủ đề tích hợp STEM cho bài học này hay không. 
      - Nếu có, bạn PHẢI TẠO MỘT HOẠT ĐỘNG HOÀN TOÀN MỚI ở CUỐI CÙNG file giáo án với tiêu đề dạng: "Hoạt động [số thứ tự]. Vận dụng STEM - [Tên chủ đề từ PPCT] (Thời gian)".
      - Hoạt động STEM này phải tuân thủ nghiêm ngặt cấu trúc 4 phần chuẩn: 
        a) Mục tiêu: (Ghi rõ mục tiêu vận dụng toán học vào thực tiễn).
        b) Nội dung (Tình huống STEM): (Đưa ra một tình huống thực tế cụ thể, yêu cầu đóng vai, ví dụ: "Kỹ sư quy hoạch...").
        c) Sản phẩm: (Yêu cầu đầu ra rõ ràng, ví dụ: Bản vẽ, mô hình, bảng dự toán...).
        d) Tổ chức thực hiện: BẮT BUỘC TRÌNH BÀY DƯỚI DẠNG BẢNG MARKDOWN (kẻ bảng 2 cột). Cột 1 là Bước (Chuyển giao nhiệm vụ, Thực hiện nhiệm vụ, Báo cáo thảo luận, Đánh giá nhận xét tổng hợp). Cột 2 là Nội dung.
      - TOÀN BỘ HOẠT ĐỘNG STEM MỚI NÀY PHẢI ĐƯỢC BỌC TRONG THẺ [BLUE] (Màu xanh lam đậm) để phân biệt rõ ràng.
      - TUYỆT ĐỐI TÁCH BIỆT: Hoạt động STEM này KHÔNG ĐƯỢC ghép chung vào TARGET của phần Mục tiêu. Nó BẮT BUỘC PHẢI NẰM DƯỚI MỘT TARGET ĐỘC LẬP LÀ: [TARGET: END_OF_DOCUMENT]. Hệ thống sẽ tự động ghép toàn bộ khối này vào cuối file Word.

      Các yêu cầu tích hợp bổ sung:
      - Năng lực AI: ${options.ai ? 'Có' : 'Không'}
      - Giáo dục hòa nhập: ${options.inclusion ? 'Có' : 'Không'}
      - Ngoại ngữ (CLIL): ${options.language ? 'Có' : 'Không'}
      - Song ngữ Việt - Anh: ${options.bilingual ? 'Có' : 'Không'}
      - Dành cho Toán KTD (STEM): ${options.toanKTD ? 'Có' : 'Không'}

      Dựa trên nội dung file giáo án "${lessonPlanFile.name}" dưới đây ${curriculumFile ? `(và file PPCT đính kèm)` : ''} ${textbookFile ? `(kết hợp với dữ liệu Sách giáo khoa đính kèm để hiểu sâu bài học)` : ''}, hãy sinh ra nội dung.
      
      --- NỘI DUNG GIÁO ÁN ---
      ${fileTextContent}
      ------------------------
      
      Trình bày nội dung xuất ra thành một văn bản thuần túy (raw text) có sử dụng các thẻ để đánh dấu màu sắc và vị trí:
      - Bao bọc các đoạn text cần bôi màu đỏ bằng thẻ [RED]...[/RED]. (Mọi nội dung liên quan đến Năng lực số, bao gồm mục 2.3, mục tiêu hoạt động, và hành động tổ chức thực hiện ĐỀU PHẢI ĐƯỢC BỌC TRONG THẺ [RED]).
      - Bao bọc các đoạn text cần bôi màu xanh lam bằng thẻ [BLUE]...[/BLUE].
      - KHÔNG DÙNG MARKDOWN CODE BLOCK BỌC KẾT QUẢ, TRẢ VỀ RAW TEXT VỚI THẺ.

      YÊU CẦU ĐỊNH DẠNG ĐẦU RA BẮT BUỘC ĐỂ HỆ THỐNG TỰ ĐỘNG CHÈN (QUAN TRỌNG NHẤT):
      Bạn PHẢI chia nội dung thành các khối [TARGET: <từ khóa trích xuất nguyên văn>].
      Vì hệ thống máy tính sẽ tìm chính xác <từ khóa> này trong file Word gốc và chèn nội dung của bạn VÀO NGAY SAU đoạn text đó. 
      DO ĐÓ, TỪ KHÓA TRONG TARGET PHẢI ĐÁP ỨNG 3 ĐIỀU KIỆN:
      1. Là MỘT TIÊU ĐỀ MỤC hoặc MỘT CÂU VĂN TRỌN VẸN CÓ THẬT trong giáo án gốc. 
      2. KHÔNG TỰ CHẾ TỪ KHÓA. Đảm bảo copy chính xác nguyên văn một câu ở cột phải của bảng.
      3. CÁC KHỐI TARGET PHẢI ĐƯỢC SẮP XẾP ĐÚNG THEO THỨ TỰ XUẤT HIỆN TRONG GIÁO ÁN TỪ TRÊN XUỐNG DƯỚI.
      
      Ví dụ mẫu đầu ra mong muốn:
      [TARGET: b) Năng lực đặc thù] (hoặc copy đúng tiêu đề mục tiêu năng lực trong file)
      [RED]2.3. Năng lực số:
      - Tích hợp Năng lực số: 1.1.NC1b...[/RED]
      ${options.ai ? '[BLUE]2.4. Lồng ghép AI:\n- Tích hợp Lồng ghép AI...[/BLUE]' : ''}
      ${options.toanKTD ? '[RED]2.5. Lồng ghép STEM:\n- Tích hợp Lồng ghép STEM...[/RED]' : ''}

      [TARGET: a) Mục tiêu]
      [RED]- Tích hợp năng lực số: 1.1.NC1b: Biết ứng dụng phần mềm/máy tính để chuyển đổi đơn vị nhằm tính toán nhanh chóng các đặc trưng cơ bản...[/RED]

      [TARGET: GV yêu cầu HS đọc yêu cầu của HĐ1 rồi yêu cầu thảo luận nhóm đôi để trả lời các câu hỏi.]
      [RED]- GV hướng dẫn quy trình thao tác bấm máy tính cầm tay (chế độ Thống kê - Statistics) để nhập giá trị đại diện...[/RED]
      
      [TARGET: HS thực hiện cá nhân Ví dụ 1, sau đó GV mời HS trả lời Ví dụ 1.]
      [RED]- HS thực hành thao tác bấm máy tính cầm tay để tính $s^2, s$ cho Ví dụ 1 và đối chiếu kiểm tra lại kết quả...[/RED]
      
      ${options.toanKTD ? `[TARGET: END_OF_DOCUMENT]
      [BLUE]Hoạt động 4.4. Vận dụng STEM - Kỹ sư Quy hoạch: Thiết kế không gian cảnh quan tối ưu (15 phút)
      a) Mục tiêu:
      - Vận dụng kiến thức về hệ trục tọa độ trong không gian...
      b) Nội dung (Tình huống STEM): Nhà trường dự kiến cải tạo một khu đất trống...
      c) Sản phẩm: Bản vẽ mặt bằng có gắn hệ trục tọa độ Oxyz...
      d) Tổ chức thực hiện:
      | Bước | Nội dung |
      |---|---|
      | Chuyển giao nhiệm vụ | GV chia lớp thành các văn phòng "Kỹ sư quy hoạch"... |
      | Thực hiện nhiệm vụ | HS phân công trong nhóm, dùng tính chất trọng tâm... |
      | Báo cáo thảo luận | Các nhóm trình chiếu mô hình... |
      | Đánh giá, nhận xét, tổng hợp | GV nhận xét sự chính xác trong các thao tác... |[/BLUE]
      ` : ''}
      `;

      const contents: any[] = [ prompt ];
      
      const fileToInlineData = async (file: File) => {
        const base64Str = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onloadend = () => {
            const result = reader.result as string;
            resolve(result.split(',')[1]);
          };
          reader.onerror = reject;
          reader.readAsDataURL(file);
        });
        
        let mimeType = file.type;
        if (!mimeType) {
          if (file.name.endsWith('.pdf')) mimeType = 'application/pdf';
          else if (file.name.endsWith('.docx')) mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
          else if (file.name.endsWith('.doc')) mimeType = 'application/msword';
          else mimeType = 'application/octet-stream';
        }
        
        return {
          inlineData: {
            data: base64Str,
            mimeType: mimeType
          }
        };
      };

      if (curriculumFile) {
        if (curriculumFile.name.endsWith('.docx')) {
          const ppctText = await extractTextFromDocx(curriculumFile);
          prompt += `\n\n--- NỘI DUNG FILE PPCT ---\n${ppctText.substring(0, 30000)}`;
          contents[0] = prompt; // Update the prompt in contents
        } else {
          contents.push(await fileToInlineData(curriculumFile));
        }
      }
      
      if (textbookFile) {
        if (textbookFile.name.endsWith('.docx')) {
          const sgkText = await extractTextFromDocx(textbookFile);
          prompt += `\n\n--- NỘI DUNG SÁCH GIÁO KHOA ---\n${sgkText.substring(0, 30000)}`;
          contents[0] = prompt; // Update the prompt in contents
        } else {
          contents.push(await fileToInlineData(textbookFile));
        }
      }

      const response = await ai.models.generateContent({
        model: aiModel,
        contents: contents
      });

      clearInterval(progressInterval);
      setProgress(100);
      setProgressText("Hoàn tất!");

      const text = response.text;
      if (text) {
        setAiResultText(text);
        setTimeout(() => {
          setIsProcessing(false);
          setShowResult(true);
          setTimeout(() => {
            const resultEl = document.getElementById('result-area');
            if (resultEl) {
              resultEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
          }, 100);
        }, 500);
      } else {
        throw new Error("AI không trả về kết quả nào.");
      }
    } catch (error: any) {
      clearInterval(progressInterval);
      setIsProcessing(false);
      console.error("Lỗi khi gọi AI:", error);
      setErrorMsg("Có lỗi xảy ra: " + (error.message || "Vui lòng kiểm tra lại API Key và thử lại."));
    }
  };

  const handleLessonPlanChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setLessonPlanFile(e.target.files[0]);
    }
  };

  const handleCurriculumChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setCurriculumFile(e.target.files[0]);
    }
  };

  const parseContentToXML = (content: string) => {
    let inTable = false;
    let currentColor = "000000";
    let xmlParagraphs = '';
    const lines = content.split('\n').filter(l => l.trim() !== '');

    for (let line of lines) {
      if (line.includes('[RED]')) currentColor = "C00000";
      if (line.includes('[BLUE]')) currentColor = "2F5496";
      
      let cleanLine = line.replace(/\[RED\]/g, '').replace(/\[\/RED\]/g, '').replace(/\[BLUE\]/g, '').replace(/\[\/BLUE\]/g, '').trim();
      cleanLine = cleanLine.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      
      if (cleanLine.startsWith('|') && cleanLine.endsWith('|')) {
          if (cleanLine.includes('---')) continue; 
          
          if (!inTable) {
              inTable = true;
              xmlParagraphs += `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders><w:tblW w:w="5000" w:type="pct"/></w:tblPr>`;
          }
          
          xmlParagraphs += `<w:tr>`;
          const cells = cleanLine.substring(1, cleanLine.length - 1).split('|');
          for (let cell of cells) {
              let text = cell.trim();
              xmlParagraphs += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:rPr><w:color w:val="${currentColor}"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
          }
          xmlParagraphs += `</w:tr>`;
      } else {
          if (inTable) {
              inTable = false;
              xmlParagraphs += `</w:tbl>`;
          }
          xmlParagraphs += `<w:p><w:r><w:rPr><w:color w:val="${currentColor}"/></w:rPr><w:t>${cleanLine}</w:t></w:r></w:p>`;
      }
      
      if (line.includes('[/RED]') || line.includes('[/BLUE]')) currentColor = "000000";
    }
    if (inTable) xmlParagraphs += `</w:tbl>`;
    return xmlParagraphs;
  };

  const handleDownload = async () => {
    if (!lessonPlanFile || !aiResultText) {
      alert("Cần tải lên file docx và phải có kết quả AI trước khi tải xuống!");
      return;
    }

    try {
      const zip = new JSZip();
      const loadedZip = await zip.loadAsync(lessonPlanFile);
      
      let docXml = await loadedZip.file("word/document.xml")?.async("string");
      if (!docXml) {
        throw new Error("Không tìm thấy tệp document.xml trong file docx này.");
      }
      
      const parts = aiResultText.split(/\[TARGET:\s*(.+?)\]/i);
      let blocks = [];
      
      if (parts.length > 1) {
        for (let i = 1; i < parts.length; i += 2) {
          const keyword = parts[i].trim().toLowerCase();
          const content = parts[i+1].trim();
          if (!content) continue;
          
          const xmlParagraphs = parseContentToXML(content);
          
          blocks.push({ keyword, xml: xmlParagraphs });
        }
      } else {
        // Fallback for old prompt results without TARGET
        const content = aiResultText;
        const xmlParagraphs = parseContentToXML(content);
        blocks.push({ keyword: 'end', xml: xmlParagraphs });
      }

      // XML Injection Logic (Smart Distribution)
      let pArray = docXml.split("</w:p>");
      
      const normalize = (str: string) => {
          if (!str) return "";
          return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/đ/g, "d").replace(/[^a-z0-9]/g, "");
      };
      
      let pTexts = pArray.map(p => normalize(p.replace(/<[^>]+>/g, '')));
      let attachments = new Array(pArray.length).fill("");
      let lastMatchedIndex = -1;
      let leftoverXml = "";
      
      for (let b = 0; b < blocks.length; b++) {
          if (blocks[b].keyword === 'end' || blocks[b].keyword === 'END_OF_DOCUMENT') {
              leftoverXml += blocks[b].xml;
              continue;
          }
          
          let kw = normalize(blocks[b].keyword);
          if (!kw) continue;
          
          let matchIndex = -1;
          // 1. Try to find the keyword after the last matched index to maintain sequence
          for (let i = lastMatchedIndex + 1; i < pTexts.length; i++) {
              if (pTexts[i].includes(kw)) {
                  matchIndex = i;
                  break;
              }
          }
          
          // 2. Fallback: search from beginning if not found
          if (matchIndex === -1) {
              for (let i = 0; i < pTexts.length; i++) {
                  if (pTexts[i].includes(kw)) {
                      matchIndex = i;
                      break;
                  }
              }
          }
          
          if (matchIndex !== -1) {
              attachments[matchIndex] += blocks[b].xml;
              lastMatchedIndex = matchIndex;
          } else {
              // If completely lost, append to leftover
              leftoverXml += blocks[b].xml;
          }
      }
      
      let modifiedXml = "";
      for (let i = 0; i < pArray.length - 1; i++) {
          modifiedXml += pArray[i] + "</w:p>";
          if (attachments[i]) {
              modifiedXml += attachments[i];
          }
      }
      modifiedXml += pArray[pArray.length - 1];
      
      // If any leftover blocks, append at the end safely
      if (leftoverXml) {
         const lastSectPrIndex = modifiedXml.lastIndexOf('<w:sectPr');
         if (lastSectPrIndex !== -1) {
             modifiedXml = modifiedXml.slice(0, lastSectPrIndex) + leftoverXml + modifiedXml.slice(lastSectPrIndex);
         } else {
             modifiedXml = modifiedXml.replace('</w:body>', `${leftoverXml}</w:body>`);
         }
      }
      
      loadedZip.file("word/document.xml", modifiedXml);
      
      const newBlob = await loadedZip.generateAsync({ type: "blob" });
      const url = window.URL.createObjectURL(newBlob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `[Da_Chen_NLS]_${lessonPlanFile.name}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (error: any) {
      console.error("Lỗi khi sửa docx:", error);
      alert("Có lỗi khi chèn vào file Word gốc: " + error.message);
    }
  };

  const handleCopy = () => {
    if (!aiResultText) return;
    const rawAiText = aiResultText.replace(/\[RED\]/g, '').replace(/\[\/RED\]/g, '').replace(/\[BLUE\]/g, '').replace(/\[\/BLUE\]/g, '');
    navigator.clipboard.writeText(rawAiText);
    alert("Đã sao chép nội dung vào bộ nhớ tạm!");
  };

  const renderColorizedText = (text: string) => {
    const parts = [];
    let currentIndex = 0;
    const regex = /\[(RED|BLUE)\]([\s\S]*?)\[\/\1\]/g;
    let match;
    
    while ((match = regex.exec(text)) !== null) {
      if (match.index > currentIndex) {
        parts.push({ type: 'normal', content: text.substring(currentIndex, match.index) });
      }
      parts.push({ type: match[1] === 'RED' ? 'red' : 'blue', content: match[2] });
      currentIndex = regex.lastIndex;
    }
    
    if (currentIndex < text.length) {
      parts.push({ type: 'normal', content: text.substring(currentIndex) });
    }

    return (
      <div className="text-left space-y-1 bg-[#f8fafc] p-6 rounded-lg border border-slate-200 shadow-inner overflow-x-auto">
        {parts.map((part, i) => {
           let classes = "whitespace-pre-wrap ";
           if (part.type === 'red') classes += "text-red-600 font-medium";
           else if (part.type === 'blue') classes += "text-blue-600 font-medium";
           else classes += "text-slate-800";
           return <span key={i} className={classes}>{part.content}</span>;
        })}
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#f0f5fa] text-slate-800 font-sans pb-20">
      {/* Header */}
      <header className="bg-[#3b72f5] text-white py-5 px-6 shadow-md flex flex-col lg:flex-row items-center justify-between gap-6 sticky top-0 z-50">
        
        {/* Logo Section */}
        <div className="flex items-center justify-center shrink-0">
          <div className="rounded-full shadow-[0_4px_15px_rgba(0,0,0,0.2)] flex items-center justify-center shrink-0 w-20 h-20 md:w-24 md:h-24 overflow-hidden border-2 border-white/20">
            <img src="/logo-toan-thay-trong.png.png" alt="Logo Toán Thầy Trọng 3T" className="w-[105%] h-[105%] object-cover" />
          </div>
        </div>

        {/* Center Text Section */}
        <div className="flex-1 flex flex-col items-center text-center px-2">
          <h1 className="text-[11px] sm:text-[14px] md:text-[16px] lg:text-[19px] xl:text-[23px] font-black uppercase tracking-normal drop-shadow-[0_2px_4px_rgba(0,0,0,0.2)] whitespace-nowrap">
            TÍCH HỢP NĂNG LỰC SỐ + AI + STEM VÀ CÁC NĂNG LỰC KHÁC VÀO KHBD
          </h1>
          <p className="text-blue-50 text-[10px] sm:text-[12px] md:text-[13px] lg:text-[14px] font-semibold mt-1.5 drop-shadow-sm opacity-90 italic whitespace-nowrap">
            Hỗ trợ tích hợp Năng lực số toàn cấp bởi Thầy Nguyễn Đình Trọng - 3T
          </p>
        </div>

        {/* Right Buttons Section */}
        <div className="flex flex-wrap justify-center items-center gap-3 shrink-0">
          <button 
            onClick={() => setShowApiModal(true)}
            className="flex items-center gap-2 bg-white/10 hover:bg-white/20 px-4 py-2.5 rounded-xl text-sm transition-all border border-white/20 backdrop-blur-sm shadow-[0_2px_8px_rgba(0,0,0,0.1)]">
            <Key className="w-4 h-4 text-yellow-400 drop-shadow-md" />
            <span className="font-bold drop-shadow-sm">{apiKey ? 'Đã cấu hình API' : 'Lấy API key'}</span>
          </button>
          <button className="hidden sm:flex items-center gap-2 bg-transparent hover:bg-white/10 px-3 py-2.5 rounded-xl text-[14px] transition-all border border-transparent hover:border-white/20">
            <BookOpen className="w-4 h-4 opacity-90 drop-shadow-sm" />
            <span className="font-bold drop-shadow-sm">Powered by Gemini</span>
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        {/* Left Column - Main Form */}
        <div className="lg:col-span-8 space-y-6">
          
          {/* Thông tin Kế hoạch bài dạy */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200/60 p-6 relative overflow-hidden">
            <div className="absolute left-0 top-0 bottom-0 w-1 bg-blue-500 rounded-l-xl"></div>
            <h2 className="text-lg font-bold text-blue-900 mb-5 flex items-center gap-2">
              <span className="bg-blue-100 text-blue-600 w-6 h-6 rounded-full flex items-center justify-center text-xs">1</span>
              Thông tin Kế hoạch bài dạy
            </h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-2">Môn học <span className="text-red-500">*</span></label>
                <select 
                  value={subject}
                  onChange={(e) => setSubject(e.target.value)}
                  className="w-full border border-slate-300 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none appearance-none bg-white text-slate-700 font-medium">
                  <option>Toán</option>
                  <option>Vật lý</option>
                  <option>Hóa học</option>
                  <option>Sinh học</option>
                  <option>Ngữ văn</option>
                  <option>Lịch sử</option>
                  <option>Địa lý</option>
                  <option>Tin học</option>
                  <option>Ngoại ngữ</option>
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-2">Khối lớp <span className="text-red-500">*</span></label>
                <select 
                  value={grade}
                  onChange={(e) => setGrade(e.target.value)}
                  className="w-full border border-slate-300 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none appearance-none bg-white text-slate-700 font-medium">
                  <option>Lớp 12</option>
                  <option>Lớp 11</option>
                  <option>Lớp 10</option>
                  <option>Lớp 9</option>
                  <option>Lớp 8</option>
                  <option>Lớp 7</option>
                  <option>Lớp 6</option>
                  <option>Tiểu học</option>
                </select>
              </div>
            </div>
          </div>

          {/* Tài liệu đầu vào */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200/60 p-6 relative overflow-hidden">
            <div className="absolute left-0 top-0 bottom-0 w-1 bg-blue-500 rounded-l-xl"></div>
            <h2 className="text-lg font-bold text-blue-900 mb-5 flex items-center gap-2">
              <span className="bg-blue-100 text-blue-600 w-6 h-6 rounded-full flex items-center justify-center text-xs">2</span>
              Tài liệu đầu vào
            </h2>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              
              {/* Upload Giáo án */}
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-3"><span className="text-red-500">*</span> File Giáo án</label>
                {lessonPlanFile ? (
                  <div className="border border-dashed border-green-400 bg-green-50/60 rounded-xl p-6 flex flex-col items-center justify-center text-center relative hover:bg-green-100/50 transition-colors group h-48">
                    <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center mb-3 shadow-sm text-green-500 group-hover:scale-110 transition-transform mx-auto">
                       <CheckCircle2 className="w-6 h-6" />
                    </div>
                    <p className="text-green-700 font-bold mb-1 px-2 truncate w-full text-xs" title={lessonPlanFile.name}>{lessonPlanFile.name}</p>
                    <button 
                      onClick={(e) => { e.preventDefault(); setLessonPlanFile(null); }}
                      className="text-green-700/70 hover:text-green-800 font-bold mb-3 text-lg leading-none cursor-pointer p-2"
                    >
                      ×
                    </button>
                    <p className="text-[13px] text-green-600 font-medium">Đã tải lên.</p>
                  </div>
                ) : (
                  <label className="border-2 border-dashed border-blue-300 bg-blue-50/30 rounded-xl p-6 flex flex-col items-center justify-center text-center relative hover:bg-blue-50/70 transition-colors cursor-pointer group h-48 w-full block">
                    <input type="file" className="hidden" accept=".doc,.docx,.pdf" onChange={handleLessonPlanChange} />
                    <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center shadow-sm text-blue-500 group-hover:scale-110 transition-transform mx-auto">
                       <UploadCloud className="w-6 h-6" />
                    </div>
                    <p className="text-blue-800 font-bold mb-1 text-sm mt-3">Tải Giáo án</p>
                    <p className="text-xs text-blue-600/70 hover:underline">Hỗ trợ .docx, .pdf</p>
                  </label>
                )}
              </div>

              {/* Upload PPCT */}
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-3">File PPCT</label>
                {curriculumFile ? (
                  <div className="border border-dashed border-green-400 bg-green-50/60 rounded-xl p-6 flex flex-col items-center justify-center text-center relative hover:bg-green-100/50 transition-colors group h-48">
                    <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center mb-3 shadow-sm text-green-500 group-hover:scale-110 transition-transform mx-auto">
                       <CheckCircle2 className="w-6 h-6" />
                    </div>
                    <p className="text-green-700 font-bold mb-1 px-2 truncate w-full text-xs" title={curriculumFile.name}>{curriculumFile.name}</p>
                    <button 
                      onClick={(e) => { e.preventDefault(); setCurriculumFile(null); }}
                      className="text-green-700/70 hover:text-green-800 font-bold mb-3 text-lg leading-none cursor-pointer p-2"
                    >
                      ×
                    </button>
                    <p className="text-[13px] text-green-600 font-medium">Đã tải lên.</p>
                  </div>
                ) : (
                  <>
                    <label className="border-2 border-dashed border-blue-200 bg-blue-50/20 rounded-xl p-6 flex flex-col items-center justify-center text-center hover:bg-blue-50/50 transition-colors cursor-pointer group h-48 w-full block">
                      <input type="file" className="hidden" accept=".doc,.docx,.pdf" onChange={handleCurriculumChange} />
                      <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center shadow-sm text-blue-400 group-hover:scale-110 transition-transform mx-auto">
                        <UploadCloud className="w-6 h-6" />
                      </div>
                      <p className="text-blue-900 font-bold mb-1 text-sm mt-3">Tải PPCT</p>
                      <p className="text-xs text-blue-500 hover:underline">Hỗ trợ .docx, .pdf</p>
                    </label>
                  </>
                )}
              </div>
              
              {/* Upload SGK */}
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-3">File SGK</label>
                {textbookFile ? (
                  <div className="border border-dashed border-green-400 bg-green-50/60 rounded-xl p-6 flex flex-col items-center justify-center text-center relative hover:bg-green-100/50 transition-colors group h-48">
                    <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center mb-3 shadow-sm text-green-500 group-hover:scale-110 transition-transform mx-auto">
                       <CheckCircle2 className="w-6 h-6" />
                    </div>
                    <p className="text-green-700 font-bold mb-1 px-2 truncate w-full text-xs" title={textbookFile.name}>{textbookFile.name}</p>
                    <button 
                      onClick={(e) => { e.preventDefault(); setTextbookFile(null); }}
                      className="text-green-700/70 hover:text-green-800 font-bold mb-3 text-lg leading-none cursor-pointer p-2"
                    >
                      ×
                    </button>
                    <p className="text-[13px] text-green-600 font-medium">Đã tải lên.</p>
                  </div>
                ) : (
                  <>
                    <label className="border-2 border-dashed border-blue-200 bg-blue-50/20 rounded-xl p-6 flex flex-col items-center justify-center text-center hover:bg-blue-50/50 transition-colors cursor-pointer group h-48 w-full block">
                      <input type="file" className="hidden" accept=".doc,.docx,.pdf" onChange={(e) => setTextbookFile(e.target.files?.[0] || null)} />
                      <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center shadow-sm text-blue-400 group-hover:scale-110 transition-transform mx-auto">
                        <BookOpen className="w-6 h-6" />
                      </div>
                      <p className="text-blue-900 font-bold mb-1 text-sm mt-3">Tải SGK</p>
                      <p className="text-xs text-blue-500 hover:underline">Hỗ trợ .docx, .pdf</p>
                    </label>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Tùy chọn nâng cao */}
          <div className="bg-[#f5f8ff] rounded-xl shadow-sm border border-blue-100 p-6 md:p-8">
            <h2 className="text-base font-bold text-blue-800 mb-5 flex items-center gap-2">
              <SlidersHorizontal className="w-5 h-5 text-blue-600" />
              Tùy chọn nâng cao
            </h2>
            
            <div className="space-y-3.5">
              {/* Checkbox 1 */}
              <label className="flex items-start gap-4 p-4 rounded-lg border border-blue-200 bg-white/60 cursor-pointer hover:bg-blue-50/80 hover:border-blue-300 transition-all shadow-sm">
                <input type="checkbox" checked={options.ai} onChange={(e) => setOptions({...options, ai: e.target.checked})} className="mt-1 w-5 h-5 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer" />
                <div>
                  <p className="font-bold text-blue-900 text-sm">Thêm năng lực trí tuệ nhân tạo vào giáo án</p>
                  <p className="text-[13px] text-blue-600/80 mt-0.5">AI sẽ phân tích và gắn năng lực AI phù hợp vào các hoạt động dạy học (hiển thị màu xanh lam)</p>
                </div>
              </label>
              
              {/* Checkbox 2 */}
              <label className="flex items-start gap-4 p-4 rounded-lg border border-purple-200 bg-white/60 cursor-pointer hover:bg-purple-50/80 hover:border-purple-300 transition-all shadow-sm">
                <input type="checkbox" checked={options.inclusion} onChange={(e) => setOptions({...options, inclusion: e.target.checked})} className="mt-1 w-5 h-5 text-purple-600 rounded border-gray-300 focus:ring-purple-500 cursor-pointer" />
                <div>
                  <p className="font-bold text-purple-900 flex items-center gap-2 text-sm">
                    <Heart className="w-4 h-4 text-purple-500"/> Thêm năng lực cho giáo án giáo dục hòa nhập
                  </p>
                  <p className="text-[13px] text-purple-600/80 mt-0.5">Sử dụng dữ liệu giáo dục hòa nhập & giải pháp hỗ trợ học sinh khuyết tật (hiển thị màu tím)</p>
                </div>
              </label>
              
              {/* Checkbox 3 */}
              <label className="flex items-start gap-4 p-4 rounded-lg border border-green-200 bg-white/60 cursor-pointer hover:bg-green-50/80 hover:border-green-300 transition-all shadow-sm">
                <input type="checkbox" checked={options.language} onChange={(e) => setOptions({...options, language: e.target.checked})} className="mt-1 w-5 h-5 text-green-600 rounded border-gray-300 focus:ring-green-500 cursor-pointer" />
                <div>
                  <p className="font-bold text-green-900 flex items-center gap-2 text-sm">
                    <Globe className="w-4 h-4 text-green-500"/> Tích hợp năng lực ngoại ngữ
                  </p>
                  <p className="text-[13px] text-green-600/80 mt-0.5">Tích hợp thuật ngữ tiếng Anh chuyên ngành theo phương pháp CLIL (hiển thị màu xanh lục)</p>
                </div>
              </label>
              
              {/* Checkbox 4 */}
              <label className="flex items-start gap-4 p-4 rounded-lg border border-teal-200 bg-white/60 cursor-pointer hover:bg-teal-50/80 hover:border-teal-300 transition-all shadow-sm">
                <input type="checkbox" checked={options.bilingual} onChange={(e) => setOptions({...options, bilingual: e.target.checked})} className="mt-1 w-5 h-5 text-teal-600 rounded border-gray-300 focus:ring-teal-500 cursor-pointer" />
                <div>
                  <p className="font-bold text-teal-900 flex items-center gap-2 text-sm">
                    <Languages className="w-4 h-4 text-teal-500"/> Tạo song ngữ Việt - Anh một phần giáo án
                  </p>
                  <p className="text-[13px] text-teal-600/80 mt-0.5">AI ưu tiên phần phù hợp như Khởi động, trò chơi hoặc hoạt động có từ khóa tiếng Anh.</p>
                </div>
              </label>

              {/* Checkbox Toán KTD */}
              <label className="flex items-start gap-4 p-4 rounded-lg border border-red-200 bg-[#fff1f2] cursor-pointer hover:bg-red-50 hover:border-red-300 transition-all shadow-sm">
                <input type="checkbox" checked={options.toanKTD} onChange={(e) => setOptions({...options, toanKTD: e.target.checked})} className="mt-1 w-5 h-5 text-red-600 rounded border-gray-300 focus:ring-red-500 cursor-pointer" />
                <div>
                  <p className="font-bold text-[#9f1239] flex items-center gap-2 text-sm">
                    <Calculator className="w-4 h-4 text-red-700"/> Dành cho Toán KTD (Tích hợp STEM)
                  </p>
                  <p className="text-[13px] text-red-700/90 mt-1 leading-snug">AI sẽ tự động dò tìm chủ đề STEM trong PPCT và sinh ra một Hoạt động Vận dụng STEM ở cuối giáo án.</p>
                </div>
              </label>
              
              {/* Link Button */}
              <a href="#" className="flex items-start gap-4 p-4 rounded-lg border border-indigo-200 bg-[#eef2ff]/50 hover:bg-[#e0e7ff] hover:border-indigo-300 transition-all text-indigo-900 block shadow-sm group">
                <div className="flex items-center gap-2 font-bold text-sm">
                  <ExternalLink className="w-4 h-4 text-indigo-500 group-hover:scale-110 transition-transform" /> Tạo song ngữ Việt - Anh toàn bộ giáo án
                </div>
                <p className="text-[13px] text-indigo-600/80 mt-1 ml-6">Nhấn để truy cập Trợ lý Song ngữ Pro.</p>
              </a>
            </div>
          </div>

          {/* Action Area */}
          {!isProcessing && !showResult && (
            <div className="flex flex-col items-center pt-2 pb-6">
              {errorMsg && (
                <div className="w-full md:w-[80%] bg-red-50 border border-red-200 text-red-600 px-4 py-3 rounded-lg mb-6 text-[14px] text-center font-medium shadow-sm flex items-center justify-center gap-2">
                  <AlertCircle className="w-5 h-5 shrink-0" /> <span className="text-left">{errorMsg}</span>
                </div>
              )}
              <button onClick={() => setShowApiModal(true)} className="text-blue-600 text-[13px] font-medium flex items-center gap-1.5 mb-5 hover:underline bg-white px-3 py-1.5 rounded-full border border-blue-100 shadow-sm">
                <Settings className="w-3.5 h-3.5" /> {apiKey ? 'Cập nhật API Key' : 'Cấu hình API Key'}
              </button>
              
              <button 
                onClick={handleStart}
                className="w-full md:w-[80%] bg-[#2563eb] hover:bg-[#1d4ed8] text-white font-bold py-4 rounded-xl shadow-[0_8px_20px_-6px_rgba(37,99,235,0.5)] flex justify-center items-center gap-2 transition-all active:scale-[0.98] text-lg uppercase tracking-wide border border-blue-500">
                <Sparkles className="w-5 h-5 text-yellow-300" /> BẮT ĐẦU SOẠN GIÁO ÁN
              </button>
            </div>
          )}

          {/* Loading Area */}
          {isProcessing && (
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden mt-8 p-12 flex flex-col items-center">
              <div className="w-full max-w-xl mx-auto text-center flex flex-col items-center">
                <div className="bg-[#1e3a8a] text-white font-medium py-3 px-12 rounded-lg mb-12 shadow-md text-[17px]">
                  Đang xử lý...
                </div>
                
                <div className="relative w-32 h-32 mx-auto mb-10">
                  <svg className="w-full h-full transform -rotate-90" viewBox="0 0 100 100">
                    <circle cx="50" cy="50" r="45" fill="none" stroke="#f1f5f9" strokeWidth="8" />
                    <circle cx="50" cy="50" r="45" fill="none" stroke="#3b82f6" strokeWidth="8" strokeDasharray="283" strokeDashoffset={283 - (283 * progress) / 100} className="transition-all duration-300 ease-out" strokeLinecap="round" />
                  </svg>
                  <div className="absolute inset-0 flex items-center justify-center text-[28px] font-bold text-blue-600">
                    {progress}%
                  </div>
                </div>
                
                <div className="w-full max-w-md bg-slate-100 rounded-full h-2 mb-6 overflow-hidden">
                  <div className="bg-blue-500 h-2 rounded-full transition-all duration-300" style={{ width: `${progress}%` }}></div>
                </div>
                
                <h3 className="text-[#1e3a8a] font-bold text-[17px] mb-3">{progressText}</h3>
                <p className="text-slate-400 text-[13px] mb-8">Powered by Gemini AI</p>
                <p className="text-orange-500/80 text-[12px] font-medium flex items-center justify-center gap-1.5">
                  <span className="text-base leading-none">💡</span> Vui lòng không đóng trang này
                </p>
              </div>
            </div>
          )}

          {/* Result Area */}
          {showResult && (
            <div id="result-area" className="bg-[#f9fbff] rounded-xl border border-blue-100 overflow-hidden mt-8 shadow-sm animate-in fade-in slide-in-from-bottom-4 duration-500">
              <div className="p-8 pb-6 text-center flex flex-col items-center">
                <div className="w-14 h-14 rounded-full flex items-center justify-center text-green-500 mb-4 border-[1.5px] border-green-500 bg-green-50/30">
                  <Check className="w-7 h-7" strokeWidth={3} />
                </div>
                <h2 className="text-[22px] font-bold text-[#1e3a8a] mb-1.5">Phân tích giáo án thành công!</h2>
                <p className="text-slate-600 mb-5 text-[15px]">Đã tạo <span className="font-bold text-slate-800">7 phần</span> nội dung NLS để chèn vào giáo án.</p>
                
                <div className="flex flex-col gap-2.5 items-center w-full max-w-2xl mb-8">
                  <div className="text-green-600 text-[13px] px-4 py-1 rounded bg-[#f0fdf4] inline-flex items-center text-center">
                    <span className="mr-1.5 font-bold">✓</span> XML Injection: giữ nguyên <span className="font-bold ml-1 mr-1">bảng biểu, MathType/OLE, hình ảnh và cấu trúc</span>; chỉ chèn nội dung mới
                  </div>
                  
                  <div className="text-red-500 text-[13px] px-4 py-1 rounded bg-[#fef2f2] inline-flex items-center text-center">
                    <span className="mr-1.5">📌</span> Nội dung NLS: <span className="font-bold ml-1 mr-1 text-red-600">màu đỏ (Times New Roman — Toán KTD)</span> • Năng lực AI: <span className="font-bold text-[#2563eb] ml-1.5 mr-1.5">màu xanh lam</span> • Phân bố vào: Cột phải bảng d) Tổ chức thực hiện
                  </div>
                </div>
                
                <div className="flex justify-center gap-3 w-full mb-5">
                  <button onClick={handleDownload} className="bg-[#2563eb] hover:bg-[#1d4ed8] text-white font-medium text-[15px] py-2.5 px-12 rounded-lg shadow-sm flex justify-center items-center gap-2 transition-all active:scale-[0.98]">
                    <Download className="w-4 h-4" /> Tải về .docx
                  </button>
                  <button onClick={handleCopy} className="bg-white hover:bg-slate-50 text-slate-600 px-4 py-2.5 rounded-lg border border-slate-200 shadow-sm transition-colors flex items-center justify-center">
                    <Copy className="w-4 h-4" />
                  </button>
                </div>
                
                <details className="w-full max-w-3xl mt-4 bg-white rounded-lg border border-slate-200/60 overflow-hidden group">
                  <summary className="px-5 py-3 cursor-pointer text-[13px] font-medium text-slate-600 bg-slate-50/50 hover:bg-slate-100 flex items-center justify-between outline-none">
                    <span className="group-open:hidden">Xem trước nội dung ({aiResultText.split('[TARGET:').length - 1} phần) <ChevronDown className="w-4 h-4 inline ml-1 opacity-70" /></span>
                    <span className="hidden group-open:block">Thu gọn xem trước <ChevronUp className="w-4 h-4 inline ml-1 opacity-70" /></span>
                  </summary>
                  <div className="p-5 border-t border-slate-100 text-[13px] text-slate-700 max-h-96 overflow-y-auto">
                    {renderColorizedText(aiResultText)}
                  </div>
                </details>
              </div>
            </div>
          )}
          
          {/* Bottom Action Area (Only shows after result) */}
          {!isProcessing && showResult && (
            <div className={`flex flex-col items-center pt-8 pb-6`}>
              {errorMsg && (
                <div className="w-full md:w-[80%] bg-red-50 border border-red-200 text-red-600 px-4 py-3 rounded-lg mb-6 text-[14px] text-center font-medium shadow-sm flex items-center justify-center gap-2">
                  <AlertCircle className="w-5 h-5 shrink-0" /> <span className="text-left">{errorMsg}</span>
                </div>
              )}
              <button onClick={() => setShowApiModal(true)} className="text-blue-600 text-[13px] font-medium flex items-center gap-1.5 mb-5 hover:underline bg-white px-3 py-1.5 rounded-full border border-blue-100 shadow-sm">
                <Settings className="w-3.5 h-3.5" /> {apiKey ? 'Cập nhật API Key' : 'Cấu hình API Key'}
              </button>
              
              <button 
                onClick={handleStart}
                className="w-full md:w-[80%] bg-[#2563eb] hover:bg-[#1d4ed8] text-white font-bold py-4 rounded-xl shadow-[0_8px_20px_-6px_rgba(37,99,235,0.5)] flex justify-center items-center gap-2 transition-all active:scale-[0.98] text-lg uppercase tracking-wide border border-blue-500">
                <Sparkles className="w-5 h-5 text-yellow-300" /> BẮT ĐẦU SOẠN LẠI GIÁO ÁN KHÁC
              </button>
            </div>
          )}

        </div>
        
        {/* Right Column - Sidebars */}
        <div className="lg:col-span-4 space-y-6">
          
          {/* Hướng dẫn nhanh (Modern Timeline Design) */}
          <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-200/60">
            <h3 className="font-bold text-[17px] text-[#1e3a8a] mb-6 flex items-center gap-2">
              <BookOpen className="w-5 h-5 text-blue-500" /> Hướng dẫn sử dụng
            </h3>
            
            <div className="relative border-l-2 border-blue-100 ml-3 space-y-7 pb-2">
              
              {/* Step 1 */}
              <div className="relative pl-6">
                <div className="absolute w-7 h-7 bg-blue-500 rounded-full flex items-center justify-center -left-[15px] top-0 ring-4 ring-white shadow-sm">
                  <span className="text-white text-xs font-bold">1</span>
                </div>
                <h4 className="text-[14px] font-bold text-slate-800 mb-1 flex items-center gap-1.5"><Key className="w-4 h-4 text-slate-500" /> Cấu hình API Key</h4>
                <p className="text-[13.5px] text-slate-600 leading-relaxed mt-1.5 mb-2">
                  Lấy miễn phí tại <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noopener noreferrer" className="text-blue-600 font-bold hover:underline">Google AI Studio</a>.
                </p>
                <div className="space-y-2 text-[13.5px] text-slate-600">
                  <div className="flex items-start gap-2">
                    <span className="text-blue-500 font-bold mt-0.5">+</span>
                    <p>Nhấn nút <b>Create an API key</b> → <b>Create a key</b>.</p>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-blue-500 font-bold mt-0.5">+</span>
                    <p><b>Copy API key</b> vừa tạo và dán vào nút <span className="bg-slate-100 px-1.5 py-0.5 rounded text-slate-700 font-medium">Cấu hình API Key</span> (ở trên cùng bên trái).</p>
                  </div>
                </div>
              </div>

              {/* Step 2 */}
              <div className="relative pl-6">
                <div className="absolute w-7 h-7 bg-blue-500 rounded-full flex items-center justify-center -left-[15px] top-0 ring-4 ring-white shadow-sm">
                  <span className="text-white text-xs font-bold">2</span>
                </div>
                <h4 className="text-[14px] font-bold text-slate-800 mb-1 flex items-center gap-1.5"><SlidersHorizontal className="w-4 h-4 text-slate-500" /> Thiết lập cơ bản</h4>
                <p className="text-[13.5px] text-slate-600 leading-relaxed mt-1.5">
                  Chọn đúng <b>Môn học</b> và <b>Khối lớp</b> của giáo án để AI phân tích chuẩn xác ngữ cảnh sư phạm.
                </p>
              </div>

              {/* Step 3 */}
              <div className="relative pl-6">
                <div className="absolute w-7 h-7 bg-blue-500 rounded-full flex items-center justify-center -left-[15px] top-0 ring-4 ring-white shadow-sm">
                  <span className="text-white text-xs font-bold">3</span>
                </div>
                <h4 className="text-[14px] font-bold text-slate-800 mb-1 flex items-center gap-1.5"><UploadCloud className="w-4 h-4 text-slate-500" /> Tải lên tài liệu</h4>
                <div className="mt-2 space-y-2 text-[13.5px] text-slate-600">
                  <div className="flex items-start gap-2">
                    <span className="text-red-500 font-bold mt-0.5">*</span>
                    <p><b>Giáo án gốc (.docx):</b> Bắt buộc để chèn mục tiêu và kịch bản tổ chức.</p>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-blue-400 font-bold mt-0.5">+</span>
                    <p><b>Kế hoạch giáo dục/PPCT (.docx):</b> (Tùy chọn) Bổ sung để AI bám sát yêu cầu năng lực đặc thù của trường.</p>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-emerald-500 font-bold mt-0.5">+</span>
                    <p><b>Sách giáo khoa (SGK):</b> (Tùy chọn) Giúp AI hiểu sâu sắc và chi tiết hơn về nội dung bài học.</p>
                  </div>
                </div>
              </div>

              {/* Step 4 */}
              <div className="relative pl-6">
                <div className="absolute w-7 h-7 bg-blue-500 rounded-full flex items-center justify-center -left-[15px] top-0 ring-4 ring-white shadow-sm">
                  <span className="text-white text-xs font-bold">4</span>
                </div>
                <h4 className="text-[14px] font-bold text-slate-800 mb-1 flex items-center gap-1.5"><Sparkles className="w-4 h-4 text-slate-500" /> Tùy chọn nâng cao</h4>
                <p className="text-[13.5px] text-slate-600 leading-relaxed mt-1.5">
                  Tích chọn các module bổ sung (Lồng ghép AI, Toán STEM, Song ngữ) nếu cần. AI sẽ tự động sinh thêm hoạt động hoặc mục tiêu tương ứng.
                </p>
              </div>

              {/* Step 5 */}
              <div className="relative pl-6">
                <div className="absolute w-7 h-7 bg-blue-500 rounded-full flex items-center justify-center -left-[15px] top-0 ring-4 ring-white shadow-sm">
                  <span className="text-white text-xs font-bold">5</span>
                </div>
                <h4 className="text-[14px] font-bold text-slate-800 mb-1 flex items-center gap-1.5"><Download className="w-4 h-4 text-slate-500" /> Xử lý & Tải về</h4>
                <p className="text-[13.5px] text-slate-600 leading-relaxed mt-1.5">
                  Nhấn <span className="bg-blue-600 text-white px-2 py-0.5 rounded text-xs font-bold uppercase">Bắt đầu soạn giáo án</span>. Chờ hệ thống phân tích (khoảng 1-2 phút) rồi tải về file Word hoàn chỉnh.
                </p>
              </div>
            </div>
            
            {/* Notes / Troubleshooting */}
            <div className="mt-8 bg-amber-50/50 border border-amber-200/60 rounded-xl p-4">
              <h4 className="text-[13px] font-bold text-amber-800 mb-2 flex items-center gap-1.5">
                <AlertCircle className="w-4 h-4" /> Khắc phục sự cố
              </h4>
              <ul className="text-[12.5px] text-amber-700/90 space-y-1.5 pl-5 list-disc">
                <li>Nếu hệ thống báo lỗi <b>Failed to fetch</b> hoặc xử lý quá chậm, API Key của bạn có thể đã quá tải (hết lượt) hoặc hết hạn. Vui lòng tạo Key mới.</li>
                <li>Hệ thống chỉ xử lý file Word chuẩn định dạng <b>.docx</b>. Vui lòng không tải lên file ảnh hoặc .doc cũ.</li>
              </ul>
            </div>
          </div>

          {/* Miền năng lực số */}
          <div className="bg-white border border-slate-200/60 rounded-xl p-6 shadow-sm">
            <h3 className="font-bold text-blue-900 text-[15px] mb-4 uppercase tracking-wider">Miền năng lực số</h3>
            <ul className="space-y-3.5 text-sm text-slate-700 font-medium">
              <li className="flex items-center gap-3">
                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></div>
                Khai thác dữ liệu và thông tin
              </li>
              <li className="flex items-center gap-3">
                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></div>
                Giao tiếp và Hợp tác
              </li>
              <li className="flex items-center gap-3">
                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></div>
                Sáng tạo nội dung số
              </li>
              <li className="flex items-center gap-3">
                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></div>
                An toàn số
              </li>
              <li className="flex items-center gap-3">
                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></div>
                Giải quyết vấn đề
              </li>
              <li className="flex items-center gap-3">
                <div className="w-1.5 h-1.5 bg-blue-500 rounded-full shrink-0 shadow-[0_0_8px_rgba(59,130,246,0.8)]"></div>
                Ứng dụng AI
              </li>
            </ul>
          </div>
        </div>

      </main>

      {/* Footer */}
      <footer className="bg-[#1e3a8a] py-12 px-6 mt-16 border-t-[6px] border-blue-400 w-full">
        <div className="max-w-4xl mx-auto">
          <p className="text-blue-200 text-[15px] mb-8 font-bold uppercase tracking-widest text-center flex items-center justify-center gap-3">
            <span className="w-12 h-[1px] bg-blue-300/30"></span>
            Mọi thông tin chi tiết liên hệ
            <span className="w-12 h-[1px] bg-blue-300/30"></span>
          </p>
          
          <div className="bg-[#172554]/50 border border-indigo-900/50 rounded-3xl p-6 sm:p-8 shadow-2xl backdrop-blur-sm max-w-[600px] mx-auto">
            
            {/* Contact Info (Centered, 3 rows) */}
            <div className="flex flex-col gap-3.5 text-center">
              
              <div className="flex flex-col sm:flex-row items-center justify-center gap-3 p-3 rounded-2xl bg-white/5 border border-white/10 hover:bg-white/10 transition-colors">
                <div className="w-8 h-8 rounded-full bg-blue-500/20 flex items-center justify-center shrink-0">
                  <span className="text-sm font-bold text-yellow-400">Z</span>
                </div>
                <div className="flex flex-col sm:flex-row items-center sm:gap-2">
                  <span className="text-blue-100 text-[12.5px] font-semibold uppercase tracking-wider">Số điện thoại Zalo:</span>
                  <span className="text-white text-[15px] font-bold tracking-wide">0352 445 795</span>
                </div>
              </div>

              <div className="flex flex-col sm:flex-row items-center justify-center gap-3 p-3 rounded-2xl bg-white/5 border border-white/10 hover:bg-white/10 transition-colors">
                <div className="w-8 h-8 rounded-full bg-blue-500/20 flex items-center justify-center shrink-0">
                  <span className="text-sm font-bold text-blue-400">O</span>
                </div>
                <div className="flex flex-col sm:flex-row items-center sm:gap-2">
                  <span className="text-blue-100 text-[12.5px] font-semibold uppercase tracking-wider">Zalo OA Toán Thầy Trọng 3T:</span>
                  <a href="https://zalo.me/4026785664346244519" target="_blank" rel="noreferrer" className="text-blue-300 hover:text-white hover:underline text-[13.5px] font-medium break-all">
                    zalo.me/4026785664346244519
                  </a>
                </div>
              </div>

              <div className="flex flex-col sm:flex-row items-center justify-center gap-3 p-3 rounded-2xl bg-white/5 border border-white/10 hover:bg-white/10 transition-colors">
                <div className="w-8 h-8 rounded-full bg-blue-500/20 flex items-center justify-center shrink-0">
                  <span className="text-sm font-bold text-blue-400">f</span>
                </div>
                <div className="flex flex-col sm:flex-row items-center sm:gap-2">
                  <span className="text-blue-100 text-[12.5px] font-semibold uppercase tracking-wider">Facebook Cá Nhân:</span>
                  <a href="https://www.facebook.com/nguyendinhtrong87" target="_blank" rel="noreferrer" className="text-blue-300 hover:text-white hover:underline text-[13.5px] font-medium break-all">
                    facebook.com/nguyendinhtrong87
                  </a>
                </div>
              </div>

            </div>

          </div>
        </div>
      </footer>

      {/* API Key Modal */}
      {showApiModal && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[100] px-4 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-xl overflow-hidden w-full max-w-[480px] shadow-2xl scale-in-95 duration-200 flex flex-col max-h-[90vh]">
            {/* Header */}
            <div className="bg-[#2563eb] text-white px-5 py-4 flex items-center justify-between">
              <h3 className="font-bold flex items-center gap-2.5 text-[15px]">
                <Key className="w-4 h-4" /> Thiết lập Model & API Key
              </h3>
              <button onClick={() => setShowApiModal(false)} className="text-blue-100 hover:text-white transition-colors p-1">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            {/* Content */}
            <div className="p-6 overflow-y-auto">
              <label className="block text-[13px] font-medium text-slate-700 mb-3">Chọn Model AI</label>
              
              <div className="space-y-3 mb-8">
                {/* Model Option 1 */}
                <div 
                  onClick={() => setAiModel('gemini-3.6-flash')}
                  className={`border rounded-xl p-3.5 cursor-pointer transition-all flex items-center gap-3 relative ${aiModel === 'gemini-3.6-flash' ? 'border-[#2563eb] ring-1 ring-[#2563eb]/20 bg-[#eff6ff]/50' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg shrink-0 ${aiModel === 'gemini-3.6-flash' ? 'bg-[#2563eb] text-white' : 'bg-slate-100 text-slate-500'}`}>
                    <Zap className="w-4 h-4" />
                  </div>
                  <div className="flex-1 pr-6">
                    <div className="flex items-center gap-2">
                      <span className={`font-semibold text-[14px] ${aiModel === 'gemini-3.6-flash' ? 'text-[#1e3a8a]' : 'text-slate-700'}`}>Gemini 3.6 Flash</span>
                      {aiModel === 'gemini-3.6-flash' && <span className="text-[10px] bg-green-100 text-green-700 px-1.5 py-0.5 rounded font-bold">Mặc định</span>}
                    </div>
                    <p className="text-[12px] text-slate-500 mt-0.5">Mặc định, GA, mạnh cho các tác vụ đa bước</p>
                  </div>
                  {aiModel === 'gemini-3.6-flash' && <CheckCircle2 className="w-5 h-5 text-[#2563eb] absolute right-4 top-1/2 -translate-y-1/2" />}
                </div>

                {/* Model Option 2 */}
                <div 
                  onClick={() => setAiModel('gemini-3.5-flash')}
                  className={`border rounded-xl p-3.5 cursor-pointer transition-all flex items-center gap-3 relative ${aiModel === 'gemini-3.5-flash' ? 'border-[#2563eb] ring-1 ring-[#2563eb]/20 bg-[#eff6ff]/50' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg shrink-0 ${aiModel === 'gemini-3.5-flash' ? 'bg-[#2563eb] text-white' : 'bg-slate-100 text-slate-500'}`}>
                    <Zap className="w-4 h-4" />
                  </div>
                  <div className="flex-1 pr-6">
                    <div className="flex items-center gap-2">
                      <span className={`font-semibold text-[14px] ${aiModel === 'gemini-3.5-flash' ? 'text-[#1e3a8a]' : 'text-slate-700'}`}>Gemini 3.5 Flash</span>
                    </div>
                    <p className="text-[12px] text-slate-500 mt-0.5">GA, dự phòng chất lượng cao</p>
                  </div>
                  {aiModel === 'gemini-3.5-flash' && <CheckCircle2 className="w-5 h-5 text-[#2563eb] absolute right-4 top-1/2 -translate-y-1/2" />}
                </div>
                
                {/* Model Option 3 */}
                <div 
                  onClick={() => setAiModel('gemini-3.5-flash-lite')}
                  className={`border rounded-xl p-3.5 cursor-pointer transition-all flex items-center gap-3 relative ${aiModel === 'gemini-3.5-flash-lite' ? 'border-[#2563eb] ring-1 ring-[#2563eb]/20 bg-[#eff6ff]/50' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg shrink-0 ${aiModel === 'gemini-3.5-flash-lite' ? 'bg-[#2563eb] text-white' : 'bg-slate-100 text-slate-500'}`}>
                    <Sparkles className="w-4 h-4" />
                  </div>
                  <div className="flex-1 pr-6">
                    <div className="flex items-center gap-2">
                      <span className={`font-semibold text-[14px] ${aiModel === 'gemini-3.5-flash-lite' ? 'text-[#1e3a8a]' : 'text-slate-700'}`}>Gemini 3.5 Flash-Lite</span>
                    </div>
                    <p className="text-[12px] text-slate-500 mt-0.5">GA, nhanh và tiết kiệm chi phí</p>
                  </div>
                  {aiModel === 'gemini-3.5-flash-lite' && <CheckCircle2 className="w-5 h-5 text-[#2563eb] absolute right-4 top-1/2 -translate-y-1/2" />}
                </div>

                {/* Model Option 4 */}
                <div 
                  onClick={() => setAiModel('gemini-3.1-flash-lite')}
                  className={`border rounded-xl p-3.5 cursor-pointer transition-all flex items-center gap-3 relative ${aiModel === 'gemini-3.1-flash-lite' ? 'border-[#2563eb] ring-1 ring-[#2563eb]/20 bg-[#eff6ff]/50' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg shrink-0 ${aiModel === 'gemini-3.1-flash-lite' ? 'bg-[#2563eb] text-white' : 'bg-slate-100 text-slate-500'}`}>
                    <Sparkles className="w-4 h-4" />
                  </div>
                  <div className="flex-1 pr-6">
                    <div className="flex items-center gap-2">
                      <span className={`font-semibold text-[14px] ${aiModel === 'gemini-3.1-flash-lite' ? 'text-[#1e3a8a]' : 'text-slate-700'}`}>Gemini 3.1 Flash-Lite</span>
                    </div>
                    <p className="text-[12px] text-slate-500 mt-0.5">Ổn định, tương thích ngược</p>
                  </div>
                  {aiModel === 'gemini-3.1-flash-lite' && <CheckCircle2 className="w-5 h-5 text-[#2563eb] absolute right-4 top-1/2 -translate-y-1/2" />}
                </div>

                {/* Model Option 5 */}
                <div 
                  onClick={() => setAiModel('gemini-2.5-flash')}
                  className={`border rounded-xl p-3.5 cursor-pointer transition-all flex items-center gap-3 relative ${aiModel === 'gemini-2.5-flash' ? 'border-[#2563eb] ring-1 ring-[#2563eb]/20 bg-[#eff6ff]/50' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg shrink-0 ${aiModel === 'gemini-2.5-flash' ? 'bg-[#2563eb] text-white' : 'bg-slate-100 text-slate-500'}`}>
                    <Zap className="w-4 h-4" />
                  </div>
                  <div className="flex-1 pr-6">
                    <div className="flex items-center gap-2">
                      <span className={`font-semibold text-[14px] ${aiModel === 'gemini-2.5-flash' ? 'text-[#1e3a8a]' : 'text-slate-700'}`}>Gemini 2.5 Flash</span>
                    </div>
                    <p className="text-[12px] text-slate-500 mt-0.5">Dự phòng cuối chuỗi</p>
                  </div>
                  {aiModel === 'gemini-2.5-flash' && <CheckCircle2 className="w-5 h-5 text-[#2563eb] absolute right-4 top-1/2 -translate-y-1/2" />}
                </div>
              </div>

              <div>
                <p className="text-[12px] text-slate-600 mb-2">Nhập Gemini API Key của bạn. Key sẽ được lưu trên trình duyệt (localStorage).</p>
                <label className="block text-[13px] font-bold text-slate-800 mb-1.5">Gemini API Key</label>
                <input
                  type="password"
                  className="w-full border-2 border-slate-800 rounded-lg px-3 py-2.5 focus:ring-0 focus:border-blue-600 outline-none font-mono text-sm tracking-widest placeholder:tracking-normal"
                  placeholder="••••••••••••••••••••••••••••••••••••"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                />
                <p className="text-[11px] text-slate-500 mt-1.5 font-medium">Hỗ trợ cả key Google AI cũ (AIzaSy...) và key mới (AQ...).</p>
              </div>
            </div>

            {/* Footer Action */}
            <div className="p-4 border-t border-slate-100 bg-slate-50 flex justify-end">
              <button 
                onClick={() => {
                  localStorage.setItem('user_gemini_api_key', apiKey);
                  localStorage.setItem('user_gemini_model', aiModel);
                  setShowApiModal(false);
                }} 
                className="px-5 py-2.5 bg-[#2563eb] hover:bg-[#1d4ed8] text-white font-medium text-[14px] rounded-lg transition-colors shadow-sm flex items-center gap-2">
                <Save className="w-4 h-4" /> Lưu cấu hình
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
