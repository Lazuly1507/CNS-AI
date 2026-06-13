from pathlib import Path
from html import escape
import re
import unicodedata

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


ROOT = Path(__file__).resolve().parent
PAGES = ROOT / "pages"
PAGES.mkdir(exist_ok=True)


PROJECTS = [
    {
        "slug": "b1",
        "code": "B1",
        "category": "Năng lực số nền tảng",
        "title": "Quản lý tệp tin và thư mục trên Windows",
        "tagline": "Từ những thao tác cơ bản đến một quy trình quản lý dữ liệu có hệ thống.",
        "summary": "Bài thực hành giúp tôi làm chủ File Explorer và hiểu rõ vòng đời của một tệp: được tạo, đặt tên, sao chép, di chuyển, xóa và khôi phục. Điểm quan trọng không chỉ là thao tác đúng, mà còn là duy trì cấu trúc dữ liệu dễ tìm và giảm rủi ro mất tệp.",
        "metrics": [("26", "ảnh minh chứng"), ("12", "bước thực hành"), ("100%", "yêu cầu hoàn thành")],
        "objectives": [
            "Tạo và tổ chức thư mục theo quy tắc đặt tên thống nhất.",
            "Phân biệt chính xác Copy/Paste và Cut/Paste trong tình huống thực tế.",
            "Hiểu sự khác nhau giữa xóa vào Recycle Bin và xóa vĩnh viễn.",
            "Ghi lại minh chứng theo trình tự để người khác có thể kiểm tra quy trình.",
        ],
        "process": [
            ("Khởi tạo không gian", "Mở File Explorer, chọn vị trí làm việc và tạo thư mục ThucHanh_NguyenNgocHieu."),
            ("Tổ chức dữ liệu", "Tạo tệp GhiChu.txt, đổi tên thành GhiChuQuanTrong.txt và tạo thư mục con TaiLieu."),
            ("Thao tác với tệp", "Sao chép tệp để giữ bản gốc; di chuyển tệp để thay đổi vị trí lưu trữ."),
            ("Quản lý xóa và phục hồi", "Thử nghiệm Delete, Shift + Delete và Restore từ Recycle Bin."),
        ],
        "results": [
            ("Quy tắc đặt tên", "Tên tệp và thư mục rõ ràng giúp nhận diện nội dung mà không cần mở tệp."),
            ("Cấu trúc thư mục", "Phân cấp hợp lý giúp tìm lại dữ liệu nhanh hơn và hỗ trợ làm việc nhóm."),
            ("An toàn dữ liệu", "Hiểu cơ chế Recycle Bin giúp lựa chọn thao tác xóa phù hợp với mức độ rủi ro."),
        ],
        "lesson": "Bài tập tưởng đơn giản nhưng tạo nền móng cho mọi công việc số phía sau. Khi dữ liệu được tổ chức tốt, việc nghiên cứu, cộng tác và sử dụng AI đều trở nên ít lỗi hơn.",
        "images": [(f"assets/b1/image{i}.png", f"Minh chứng bước {i} trong quy trình quản lý tệp") for i in [1, 3, 5, 8, 10, 13, 15, 19, 22, 25, 26]],
        "pdf": "B1_Quan-ly-tep-va-thu-muc.pdf",
    },
    {
        "slug": "b2",
        "code": "B2",
        "category": "Nghiên cứu học thuật",
        "title": "Tìm kiếm và đánh giá thông tin học thuật",
        "tagline": "Không chỉ tìm được tài liệu, mà còn biết tài liệu nào xứng đáng để tin.",
        "summary": "Tôi nghiên cứu tác động của AI tạo sinh đến năng suất lao động trong nhóm ngành dịch vụ. Quy trình đi từ xác định câu hỏi, thiết kế từ khóa, sàng lọc nguồn đến đánh giá phương pháp và tổng hợp bằng chứng.",
        "metrics": [("10", "nguồn đã kiểm chứng"), ("6", "bài báo khoa học"), ("5", "tiêu chí đánh giá")],
        "objectives": [
            "Thu thập tối thiểu 10 tài liệu từ nhiều nhóm nguồn đáng tin cậy.",
            "Đánh giá nguồn theo tác giả, nhà xuất bản, phương pháp, trích dẫn và tính cập nhật.",
            "Tổng hợp kết quả thay vì chỉ liệt kê tài liệu.",
            "Trình bày danh mục tham khảo theo định dạng Harvard.",
        ],
        "process": [
            ("Đặt câu hỏi", "AI tạo sinh ảnh hưởng thế nào đến năng suất, chất lượng công việc và nhu cầu kỹ năng?"),
            ("Thiết kế truy vấn", "Kết hợp các từ khóa generative AI, productivity, service industry và labour market."),
            ("Sàng lọc", "Ưu tiên nghiên cứu thực nghiệm, tạp chí phản biện và báo cáo chính thức giai đoạn 2022–2025."),
            ("Đánh giá & tổng hợp", "So sánh phương pháp, kết quả, hạn chế và mức độ có thể khái quát của từng nguồn."),
        ],
        "results": [
            ("Năng suất", "AI tạo sinh cải thiện rõ nhất các tác vụ có cấu trúc như viết, hỗ trợ khách hàng và lập trình."),
            ("Phân bố lợi ích", "Người ít kinh nghiệm thường hưởng lợi nhiều hơn, nhưng hiệu quả tổ chức phụ thuộc quản trị và dữ liệu."),
            ("Rủi ro", "Tốc độ tăng không bảo đảm chất lượng; kiểm chứng đầu ra vẫn là điều kiện bắt buộc."),
        ],
        "lesson": "Tôi thay đổi cách nhìn về việc “tìm tài liệu”: kết quả đứng đầu không mặc nhiên là nguồn tốt nhất. Một tài liệu đáng tin cần có phương pháp phù hợp, bối cảnh rõ và giới hạn được thừa nhận.",
        "table": {
            "headers": ["Nguồn tiêu biểu", "Loại nguồn", "Kết luận chính", "Độ tin cậy"],
            "rows": [
                ["Noy & Zhang (2023)", "Science", "Giảm thời gian, tăng chất lượng nhiệm vụ viết", "Rất cao"],
                ["Brynjolfsson et al. (2025)", "QJE", "Năng suất hỗ trợ khách hàng tăng trung bình 15%", "Rất cao"],
                ["Kassa & Worku (2025)", "Journal of Open Innovation", "Năng suất nhân viên là biến trung gian quan trọng", "Cao"],
                ["Salari et al. (2025)", "Systematic review", "GenAI vừa tạo vai trò mới vừa thay đổi nhu cầu kỹ năng", "Rất cao"],
                ["Teutloff et al. (2025)", "J. Economic Behavior", "Nhu cầu một số kỹ năng freelance giảm mạnh", "Rất cao"],
            ],
        },
        "pdf": "B2_Tim-kiem-va-danh-gia-thong-tin-hoc-thuat.pdf",
    },
    {
        "slug": "b3",
        "code": "B3",
        "category": "Prompt Engineering",
        "title": "Tối ưu hóa prompt trong học tập",
        "tagline": "Một câu hỏi tốt không làm AI thông minh hơn, nhưng giúp AI hiểu đúng điều người học cần.",
        "summary": "Ba tác vụ học tập được thử nghiệm ở ba cấp độ prompt: cơ bản, cải tiến và nâng cao. Kết quả cho thấy vai trò, ngữ cảnh, nguồn dữ liệu và cấu trúc đầu ra quyết định trực tiếp đến khả năng sử dụng câu trả lời.",
        "metrics": [("9", "prompt hoàn chỉnh"), ("3", "tác vụ học tập"), ("6", "ảnh thử nghiệm")],
        "objectives": [
            "Thiết kế ba cấp độ prompt có khác biệt rõ ràng cho mỗi tác vụ.",
            "Thử nghiệm trên giải thích SVD, tóm tắt Java Socket và tạo câu hỏi ôn tập.",
            "So sánh đầu ra theo độ chính xác, cấu trúc, tính ứng dụng và mức bám yêu cầu.",
            "Rút ra nguyên tắc viết prompt có thể tái sử dụng.",
        ],
        "process": [
            ("Prompt cơ bản", "Nêu yêu cầu ngắn, phù hợp khám phá nhanh nhưng đầu ra thường chung chung."),
            ("Prompt cải tiến", "Bổ sung đối tượng, nội dung cần có, giới hạn và định dạng đầu ra."),
            ("Prompt nâng cao", "Kết hợp role prompting, grounding, delimiters, ma trận yêu cầu và tự kiểm tra."),
            ("Đánh giá", "Chấm đầu ra theo thang 1–5 và quan sát khả năng dùng trực tiếp trong học tập."),
        ],
        "results": [
            ("Role + Context", "Vai trò và đối tượng giúp AI chọn đúng độ sâu và cách diễn đạt."),
            ("Output schema", "Quy định bảng, workflow hay phần đề/đáp án giúp giảm đáng kể thời gian biên tập."),
            ("Grounding", "Yêu cầu bám tài liệu và đánh dấu phần chưa nêu giúp hạn chế thông tin ngoài nguồn."),
        ],
        "lesson": "Prompt dài hơn không tự động tốt hơn. Prompt hiệu quả là prompt có mục đích: mỗi ràng buộc đều phục vụ một tiêu chí đầu ra cụ thể.",
        "table": {
            "headers": ["Tác vụ", "Cơ bản", "Cải tiến", "Nâng cao"],
            "rows": [
                ["Giải thích SVD", "Định nghĩa ngắn", "Có công thức và ứng dụng", "Trực giác → công thức → tự kiểm tra"],
                ["Tóm tắt Java Socket", "Tóm tắt chung", "Bốn nhóm nội dung", "Bám nguồn, workflow và lỗi hiểu sai"],
                ["Quiz Java Socket", "10 câu hỏi", "Đa dạng dạng câu", "Ma trận Bloom, code và tình huống"],
            ],
        },
        "images": [(f"assets/b3/image{i}.png", f"Ảnh thử nghiệm prompt và đầu ra số {i}") for i in range(1, 7)],
        "pdf": "B3_Toi-uu-prompt-trong-hoc-tap.pdf",
    },
    {
        "slug": "b4",
        "code": "B4",
        "category": "Cộng tác trực tuyến",
        "title": "Điều phối dự án Auction Online System",
        "tagline": "Công cụ không thay thế teamwork; công cụ làm cho trách nhiệm và tiến độ trở nên nhìn thấy được.",
        "summary": "Trong vai trò Core Backend, tôi sử dụng Trello, Google Docs, Google Drive, Discord và GitHub để quản lý nhiệm vụ, đóng góp tài liệu, tổ chức tài nguyên và hỗ trợ kỹ thuật cho dự án đấu giá trực tuyến.",
        "metrics": [("5", "công cụ phối hợp"), ("9", "ảnh minh chứng"), ("3", "thách thức đã xử lý")],
        "objectives": [
            "Quản lý nhiệm vụ cá nhân có trạng thái, mô tả, checklist và deadline.",
            "Minh chứng đóng góp trực tiếp trên tài liệu cộng tác.",
            "Tổ chức tệp theo cấu trúc nhiều cấp và phân quyền phù hợp.",
            "Duy trì giao tiếp chủ động trong quá trình phát triển backend.",
        ],
        "process": [
            ("Trello", "Dùng Backlog, To Do, In Progress, Testing và Done; kết hợp label và checklist."),
            ("Google Docs", "Soạn đặc tả luồng dữ liệu, hướng dẫn database và trao đổi bằng Comments."),
            ("Google Drive", "Tổ chức tài liệu theo ba cấp và áp dụng quy tắc đặt tên thống nhất."),
            ("Discord & GitHub", "Giải đáp lỗi Socket, chia sẻ code và quản lý phiên bản nguồn."),
        ],
        "results": [
            ("Xung đột phiên bản", "Giải quyết bằng lịch sử chỉnh sửa, Suggesting và quy tắc thông báo trước khi sửa."),
            ("Trôi thông tin", "Dùng thread và pinned message để lưu quyết định kỹ thuật quan trọng."),
            ("Quá tải nhiệm vụ", "Chia task thành sub-task dưới hai giờ và ưu tiên bằng deadline/label."),
        ],
        "lesson": "Hiệu quả cao nhất xuất hiện khi mỗi loại thông tin được đặt đúng công cụ: task ở Trello, tài liệu ở Docs, tệp ở Drive, trao đổi nhanh ở Discord và mã nguồn ở GitHub.",
        "images": [(f"assets/b4/image{i}.png", f"Minh chứng cộng tác trực tuyến số {i}") for i in range(1, 10)],
        "pdf": "B4_Cong-tac-truc-tuyen.pdf",
    },
    {
        "slug": "b5",
        "code": "B5",
        "category": "Sáng tạo với AI",
        "title": "Chiến dịch Cà phê xanh · Sống lành",
        "tagline": "AI tạo phương án; con người tạo ra bản sắc và chịu trách nhiệm về sản phẩm cuối.",
        "summary": "Dự án phối hợp AI tạo văn bản, hình ảnh và thiết kế để xây dựng bộ nội dung truyền thông cho một quán cà phê bền vững. Đầu ra AI được xem như nguyên liệu, sau đó được chọn lọc, sửa lỗi và thiết kế lại.",
        "metrics": [("3", "công cụ AI"), ("4", "phiên bản/minh chứng"), ("55%", "đóng góp cá nhân")],
        "objectives": [
            "Tạo thông điệp nhất quán về cà phê hữu cơ và tái sử dụng bã cà phê.",
            "Ghi lại prompt, đầu ra và quyết định chỉnh sửa cho từng công cụ.",
            "So sánh điểm mạnh, hạn chế và vai trò phù hợp của ba công cụ.",
            "Hoàn thiện sản phẩm có đóng góp sáng tạo cá nhân rõ ràng.",
        ],
        "process": [
            ("Google Gemini", "Phát triển big idea, đối tượng mục tiêu và kế hoạch nội dung bốn tuần."),
            ("Gemini Image Generator", "Tạo key visual mang không khí tự nhiên với tông xanh–nâu."),
            ("Canva AI", "Tạo bộ icon và bố cục nháp; phát hiện lỗi chữ sinh tự động."),
            ("Biên tập cá nhân", "Viết lại nội dung, sửa lỗi, tổ chức bốn bước và hoàn thiện infographic."),
        ],
        "results": [
            ("Văn bản", "Gemini tạo khung nhanh nhưng cần thu hẹp chủ đề và thay nội dung chung chung."),
            ("Hình ảnh", "Key visual đúng không khí nhưng chưa thể hiện nhận diện và thông điệp cụ thể."),
            ("Thiết kế", "Canva AI hữu ích để phác thảo; chữ và logic thông tin phải được sửa thủ công."),
        ],
        "lesson": "Tốc độ tạo sinh chỉ thực sự có giá trị khi đi kèm khả năng chọn lọc. Phần sáng tạo quan trọng nhất nằm ở quyết định giữ gì, bỏ gì và kết nối các đầu ra thành một câu chuyện nhất quán.",
        "images": [
            ("assets/b5/final-infographic.png", "Infographic cuối của chiến dịch"),
            ("assets/b5/image1.png", "Gemini xây dựng kế hoạch nội dung"),
            ("assets/b5/image2.png", "Key visual do AI tạo hình ảnh"),
            ("assets/b5/image3.png", "Phiên bản trung gian trên Canva AI"),
        ],
        "pdf": "B5_Sang-tao-noi-dung-voi-AI.pdf",
    },
    {
        "slug": "b6",
        "code": "B6",
        "category": "Đạo đức số",
        "title": "Sử dụng AI có trách nhiệm trong học thuật",
        "tagline": "Minh bạch không làm giảm giá trị bài làm; minh bạch chứng minh người học đang làm chủ công cụ.",
        "summary": "Bài tập kết hợp phân tích chính sách, thực hành một nhiệm vụ với AI, đánh giá vấn đề đạo đức và xây dựng bộ nguyên tắc cá nhân H.E.R.O.I.C để duy trì tính chính trực học thuật.",
        "metrics": [("6", "nguyên tắc cá nhân"), ("3", "vấn đề đạo đức"), ("4", "bước dùng AI an toàn")],
        "objectives": [
            "Phân tích định hướng công khai về AI và liêm chính học thuật.",
            "Ghi lại prompt, đầu ra, cách đánh giá và chỉnh sửa trong một nhiệm vụ thực tế.",
            "Phân biệt hỗ trợ hợp lý với gian lận học thuật.",
            "Xây dựng nguyên tắc có thể áp dụng lâu dài trong học tập.",
        ],
        "process": [
            ("Hỏi có mục đích", "Chỉ sử dụng AI cho nhiệm vụ được phép và xác định rõ mục tiêu trước khi hỏi."),
            ("Kiểm chứng", "Đối chiếu số liệu, trích dẫn và lập luận với nguồn gốc đáng tin cậy."),
            ("Cá nhân hóa", "Viết lại bằng tư duy, trải nghiệm và ngôn ngữ của bản thân."),
            ("Minh bạch", "Khai báo công cụ, mục đích và phạm vi AI đã hỗ trợ."),
        ],
        "results": [
            ("Hỗ trợ vs gian lận", "AI là co-pilot khi người học vẫn tư duy, kiểm chứng và chịu trách nhiệm."),
            ("Sở hữu trí tuệ", "Không thể xem đầu ra AI là nguồn gốc; cần tìm và trích dẫn tài liệu gốc."),
            ("Phát triển kỹ năng", "Lạm dụng AI làm yếu năng lực cốt lõi; dùng đúng cách tạo thời gian cho tư duy bậc cao."),
        ],
        "lesson": "Nguyên tắc cốt lõi của tôi là Human-in-the-loop: con người luôn là người quyết định, kiểm chứng và chịu trách nhiệm cuối cùng.",
        "special": "heroic",
        "images": [("assets/b6/image1.png", "Infographic sử dụng AI có trách nhiệm trong học thuật")],
        "pdf": "B6_Su-dung-AI-co-trach-nhiem.pdf",
    },
    {
        "slug": "b7",
        "code": "B7",
        "category": "AI cho nghiên cứu",
        "title": "Tổng quan Graphene trong pin Lithium-ion",
        "tagline": "AI giúp đọc nhanh hơn; lựa chọn đúng tài liệu vẫn cần phán đoán khoa học của con người.",
        "summary": "Tôi sử dụng Elicit để tìm kiếm và tổng hợp năm nghiên cứu thực nghiệm trực tiếp về graphene trong pin Lithium-ion, sau đó kiểm tra DOI, phương pháp, kết quả, vai trò của graphene và hạn chế.",
        "metrics": [("5", "bài báo thực nghiệm"), ("5", "DOI đã kiểm tra"), ("3", "cơ chế cải thiện")],
        "objectives": [
            "Đặt câu hỏi nghiên cứu đủ hẹp và có khả năng tổng hợp.",
            "Sàng lọc đúng các nghiên cứu trực tiếp về graphene trong pin Lithium-ion.",
            "Trích xuất hơn bốn nhóm thông tin từ mỗi bài báo.",
            "Đưa ra nhận xét tổng hợp thay vì chỉ mô tả từng nghiên cứu riêng lẻ.",
        ],
        "process": [
            ("Tìm kiếm trên Elicit", "Dùng các truy vấn graphene lithium-ion battery anode và graphene composite."),
            ("Sàng lọc", "Loại nghiên cứu về siêu tụ điện hoặc vật liệu không tập trung vào graphene."),
            ("Trích xuất", "Ghi nhận phương pháp, vật liệu, kết quả, vai trò graphene và hạn chế."),
            ("Kiểm chứng", "Đối chiếu tiêu đề, tạp chí và DOI với nguồn xuất bản."),
        ],
        "results": [
            ("Lưu trữ lithium", "Graphene nanosheet cung cấp nhiều vị trí lưu trữ và có thể vượt dung lượng graphite."),
            ("Độ dẫn điện", "Mạng graphene giúp vận chuyển electron và cải thiện hiệu suất ở tốc độ cao."),
            ("Ổn định cấu trúc", "Graphene hạn chế kết tụ và biến đổi thể tích của vật liệu điện cực."),
        ],
        "lesson": "Elicit rất hiệu quả để tạo bản đồ tài liệu ban đầu, nhưng kết quả tìm kiếm có thể chứa nguồn gần chủ đề mà không thực sự trả lời câu hỏi. Bước sàng lọc và kiểm tra DOI là không thể bỏ qua.",
        "table": {
            "headers": ["Nghiên cứu", "Vật liệu", "Kết quả nổi bật", "Hạn chế"],
            "rows": [
                ["Yoo et al. (2008)", "GNS + CNT/C60", "540–784 mAh g⁻¹", "Quy mô sản xuất"],
                ["Wang et al. (2009)", "TiO₂–graphene", "Tăng hơn hai lần ở tốc độ cao", "Chế tạo phức tạp"],
                ["Zhou et al. (2010)", "Graphene–Fe₃O₄", "Ổn định chu kỳ và tốc độ cao", "Phụ thuộc cấu trúc"],
                ["Vargas et al. (2013)", "GNS/LNMO full-cell", "Điện áp trung bình 3,75 V", "Dung lượng full-cell"],
                ["Jiao et al. (2015)", "Graphitized graphene", "Cải thiện chu kỳ và rate", "Chi phí vật liệu"],
            ],
        },
        "images": [
            ("assets/b7/image1.png", "Kết quả tìm kiếm và trích xuất trên Elicit"),
            ("assets/b7/image2.png", "Các nghiên cứu khác trong bảng kết quả Elicit"),
        ],
        "pdf": "B7_Tong-hop-tai-lieu-khoa-hoc-bang-AI.pdf",
    },
]

INLINE_EVIDENCE = {
    "B3": {
        "III. KẾT QUẢ THỬ NGHIỆM VÀ SO SÁNH": [
            ("assets/b3/image1.png", "SVD: prompt cơ bản và đầu ra"),
            ("assets/b3/image2.png", "SVD: prompt nâng cao và đầu ra"),
            ("assets/b3/image3.png", "Java Socket: prompt tóm tắt cơ bản"),
            ("assets/b3/image4.png", "Java Socket: prompt tóm tắt nâng cao"),
            ("assets/b3/image5.png", "Quiz Java Socket: prompt cơ bản"),
            ("assets/b3/image6.png", "Quiz Java Socket: prompt nâng cao"),
        ],
    },
    "B4": {
        "1. Quản lý tác vụ với Trello": [
            ("assets/b4/image1.png", "Bảng Trello sau khi các nhiệm vụ chính được hoàn thành"),
            ("assets/b4/image2.png", "Lịch sử hoạt động Trello thể hiện quá trình cập nhật tiến độ"),
            ("assets/b4/image7.png", "Checklist được sử dụng để chia nhỏ nhiệm vụ cá nhân"),
        ],
        "2. Soạn thảo tài liệu với Google Docs": [
            ("assets/b4/image3.png", "Version History trên Google Docs thể hiện phần đóng góp cá nhân"),
            ("assets/b4/image8.png", "Quá trình tạo và hoàn thiện báo cáo cộng tác"),
            ("assets/b4/image9.png", "Lịch sử các phiên chỉnh sửa trên Google Docs"),
        ],
        "3. Quản lý tài nguyên trên Google Drive": [
            ("assets/b4/image4.png", "Cấu trúc thư mục cấp đầu trên Google Drive"),
            ("assets/b4/image5.png", "Cấu trúc thư mục con dùng để tổ chức tài nguyên dự án"),
            ("assets/b4/image6.png", "Thiết lập quyền truy cập phù hợp cho một tệp dùng chung"),
        ],
    },
    "B5": {
        "2.1. Google Gemini": [("assets/b5/image1.png", "Prompt và kế hoạch nội dung do Google Gemini đề xuất")],
        "2.2. Gemini Image Generator": [("assets/b5/image2.png", "Key visual được tạo từ prompt mô tả")],
        "2.3. Canva AI": [("assets/b5/image3.png", "Phiên bản trung gian trên Canva AI còn lỗi chữ và bố cục")],
        "3.1. Sản phẩm cuối": [("assets/b5/final-infographic.png", "Infographic cuối sau khi sửa chữ và hoàn thiện bố cục")],
    },
    "B6": {
        "PHẦN 5: Ý TƯỞNG & THIẾT KẾ INFOGRAPHIC": [
            ("assets/b6/image1.png", "Infographic sử dụng AI có trách nhiệm trong học thuật"),
        ],
    },
    "B7": {
        "5. MINH CHỨNG SỬ DỤNG ELICIT": [
            ("assets/b7/image1.png", "Kết quả tìm kiếm trên Elicit với từ khóa về graphene và pin lithium-ion"),
            ("assets/b7/image2.png", "Bảng Elicit trích xuất phương pháp và kết quả của các bài báo được chọn"),
        ],
    },
}

SKIP_REPORT_HEADINGS = {
    "B3": ("PHỤ LỤC B – MINH CHỨNG THỬ NGHIỆM",),
    "B4": ("VI. PHỤ LỤC MINH CHỨNG",),
}


def header(prefix="", active=""):
    links = [
        ("home", f"{prefix}index.html", "Trang chủ"),
        ("projects", f"{prefix}index.html#projects", "Bài tập"),
        ("reflection", f"{prefix}pages/tong-ket.html", "Tổng kết"),
    ]
    nav = "".join(f'<a class="{"active" if key == active else ""}" href="{url}">{label}</a>' for key, url, label in links)
    return f"""
    <a class="skip-link" href="#main">Đi đến nội dung chính</a>
    <div class="scroll-progress" aria-hidden="true"><span></span></div>
    <div class="page-transition" aria-hidden="true"><div><span>NH</span><small>Digital AI Portfolio</small></div></div>
    <div class="cursor-glow" aria-hidden="true"></div>
    <canvas id="particle-canvas" aria-hidden="true"></canvas>
    <header class="site-header">
      <div class="shell nav-wrap">
        <a class="brand" href="{prefix}index.html" aria-label="Trang chủ">
          <img src="{prefix}assets/logo.png" alt="UET">
          <span><strong>Digital AI Portfolio</strong><small>Nguyễn Ngọc Hiếu · VNU-UET</small></span>
        </a>
        <button class="nav-toggle" aria-label="Mở menu" aria-expanded="false"><span></span><span></span></button>
        <nav class="site-nav" aria-label="Điều hướng chính">{nav}</nav>
      </div>
    </header>"""


def footer(prefix=""):
    return f"""
    <footer class="site-footer">
      <div class="shell footer-grid">
        <div><span class="footer-mark">NH</span><p>Hành trình học tập về năng lực số và trí tuệ nhân tạo.</p></div>
        <div class="footer-links">
          <a href="{prefix}index.html">Trang chủ</a>
          <a href="{prefix}index.html#projects">Bảy bài tập</a>
          <a href="{prefix}pages/tong-ket.html">Tổng kết học phần</a>
          <a href="mailto:25021764@gmail.com">25021764@gmail.com</a>
        </div>
      </div>
      <div class="shell footer-bottom"><span>VNU-UET · VNU1001_E252012 · UET.A12</span><span>Thiết kế & nội dung bởi Nguyễn Ngọc Hiếu</span></div>
    </footer>
    <button class="back-top" aria-label="Lên đầu trang">↑</button>
    <dialog class="lightbox"><button class="lightbox-close" aria-label="Đóng">×</button><img src="" alt=""><p></p></dialog>
    <script src="{prefix}portfolio.js?v=20260613-2"></script>"""


def doc(title, body, prefix="", active="", description="Portfolio học tập về năng lực số và AI."):
    return f"""<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="{escape(description)}">
  <meta name="theme-color" content="#07101f">
  <title>{escape(title)} | Nguyễn Ngọc Hiếu</title>
  <link rel="icon" href="{prefix}assets/logo.png">
  <link rel="stylesheet" href="{prefix}portfolio.css?v=20260613-2">
</head>
<body>
{header(prefix, active)}
<main id="main">{body}</main>
{footer(prefix)}
</body>
</html>"""


def metrics_html(metrics):
    return "".join(f'<div><strong class="counter" data-target="{escape(value)}">{escape(value)}</strong><span>{escape(label)}</span></div>' for value, label in metrics)


def cards(items, numbered=False):
    result = []
    for index, (title, text) in enumerate(items, 1):
        number = f"<span>{index:02d}</span>" if numbered else '<i aria-hidden="true"></i>'
        result.append(f'<article class="insight-card tilt reveal">{number}<h3>{escape(title)}</h3><p>{escape(text)}</p></article>')
    return "".join(result)


def table_html(table):
    if not table:
        return ""
    head = "".join(f"<th>{escape(x)}</th>" for x in table["headers"])
    rows = "".join("<tr>" + "".join(f"<td>{escape(x)}</td>" for x in row) + "</tr>" for row in table["rows"])
    return f'<div class="data-table reveal"><table><thead><tr>{head}</tr></thead><tbody>{rows}</tbody></table></div>'


def iter_docx_blocks(document):
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def anchor_id(text, used):
    normalized = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode().lower()
    slug = re.sub(r"[^a-z0-9]+", "-", normalized).strip("-") or "muc"
    count = used.get(slug, 0) + 1
    used[slug] = count
    return slug if count == 1 else f"{slug}-{count}"


def report_heading(text, style_name, base_style_level=None):
    if style_name.startswith("Heading"):
        match = re.search(r"(\d+)", style_name)
        level = int(match.group(1)) if match else 2
        if base_style_level is not None:
            return min(3, 2 + max(0, level - base_style_level))
        return 2 if level <= 3 else 3
    clean = text.strip()
    if re.match(r"^PHẦN\s+\d+", clean) and clean.upper() == clean:
        return 2
    if re.match(r"^([IVX]+\.\s+|PHỤ LỤC)", clean, re.I):
        return 2
    if len(clean) < 100 and re.match(r"^(TÀI LIỆU THAM KHẢO|KẾT LUẬN)", clean, re.I):
        return 2
    if re.match(r"^\d+\.\d+\.?\s+", clean):
        return 3
    if re.match(r"^\d+\.\s+", clean):
        return 2 if clean.upper() == clean else 3
    return None


def report_table_html(table):
    rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
    if not rows or not any(any(cell for cell in row) for row in rows):
        return ""
    flat = " ".join(cell for row in rows for cell in row)
    if len(rows) <= 4 and len(rows[0]) == 2 and ("Họ và tên" in flat or "Sinh viên" in flat):
        return ""
    nonempty = [cell for row in rows for cell in row if cell]
    if nonempty and all(cell.lower().startswith(("hình ", "minh chứng")) for cell in nonempty):
        return ""
    if len(rows) == 1 and len(rows[0]) == 1:
        parts = [part.strip() for part in rows[0][0].splitlines() if part.strip()]
        if not parts:
            return ""
        title = f"<strong>{escape(parts[0])}</strong>" if len(parts) > 1 else ""
        text = " ".join(parts[1:] if len(parts) > 1 else parts)
        return f'<aside class="report-callout reveal">{title}<p>{escape(text)}</p></aside>'
    head = "".join(f"<th>{escape(cell)}</th>" for cell in rows[0])
    body = "".join("<tr>" + "".join(f"<td>{escape(cell)}</td>" for cell in row) + "</tr>" for row in rows[1:])
    return f'<div class="report-table reveal"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


def b1_tutorial():
    stages = [
        (
            "01",
            "Chuẩn bị không gian làm việc",
            "Tạo một vị trí thực hành riêng để mọi thao tác phía sau có cấu trúc rõ ràng và dễ kiểm tra.",
            [
                ("Mở File Explorer", "Dùng Windows + E để mở nhanh trình quản lý tệp.", [(1, "Mở File Explorer trên Windows")]),
                ("Chọn vị trí lưu", "Đi đến ổ dữ liệu hoặc thư mục Documents, tránh thao tác nhầm trong thư mục hệ thống.", [(2, "Truy cập vị trí lưu thư mục thực hành")]),
                ("Tạo thư mục thực hành", "Tạo thư mục mới và đặt tên theo quy tắc ThucHanh_HoTen.", [(3, "Chọn lệnh New → Folder"), (4, "Đặt tên thư mục thực hành")]),
                ("Mở thư mục vừa tạo", "Kiểm tra thanh địa chỉ để chắc chắn mọi nội dung tiếp theo được tạo đúng vị trí.", [(5, "Mở thư mục thực hành vừa tạo")]),
            ],
        ),
        (
            "02",
            "Tạo và tổ chức dữ liệu",
            "Xây dựng cấu trúc nhỏ gồm tệp ghi chú và thư mục con, đồng thời thực hành quy tắc đặt tên dễ hiểu.",
            [
                ("Tạo tệp văn bản", "Chọn New → Text Document, sau đó đặt tên ban đầu là GhiChu.txt.", [(6, "Chọn tạo Text Document"), (7, "Đặt tên tệp GhiChu.txt")]),
                ("Đổi tên có ý nghĩa", "Đổi GhiChu.txt thành GhiChuQuanTrong.txt để tên tệp phản ánh đúng nội dung.", [(8, "Chọn thao tác Rename"), (9, "Kết quả sau khi đổi tên")]),
                ("Tạo thư mục con", "Tạo thư mục TaiLieu để tách tài liệu khỏi vùng làm việc chính.", [(10, "Chọn tạo thư mục con"), (11, "Cấu trúc sau khi có thư mục TaiLieu")]),
            ],
        ),
        (
            "03",
            "Phân biệt sao chép và di chuyển",
            "Quan sát trực tiếp sự khác nhau: Copy tạo thêm một bản, còn Cut thay đổi vị trí của bản gốc.",
            [
                ("Copy & Paste", "Sao chép GhiChuQuanTrong.txt vào TaiLieu. Sau thao tác, tệp tồn tại ở cả vị trí cũ và mới.", [(12, "Chọn Copy đối với tệp"), (13, "Paste vào thư mục TaiLieu"), (14, "Bản sao tại vị trí mới"), (15, "Bản gốc vẫn ở vị trí cũ")]),
                ("Cut & Paste", "Tạo DiChuyen.txt rồi di chuyển vào TaiLieu. Sau thao tác, tệp chỉ còn ở vị trí mới.", [(16, "Tạo tệp DiChuyen.txt"), (17, "Chọn Cut đối với tệp"), (18, "Paste vào thư mục TaiLieu"), (19, "Tệp tại vị trí mới"), (20, "Vị trí cũ không còn tệp")]),
            ],
        ),
        (
            "04",
            "Xóa an toàn và khôi phục",
            "Thử cả xóa tạm thời, xóa vĩnh viễn và phục hồi để hiểu vòng đời dữ liệu trên Windows.",
            [
                ("Xóa vào Recycle Bin", "Dùng Delete với GhiChuQuanTrong.txt. Tệp biến mất khỏi thư mục nhưng vẫn có thể phục hồi.", [(21, "Trạng thái trước khi Delete"), (22, "Trạng thái sau khi Delete")]),
                ("Xóa vĩnh viễn", "Dùng Shift + Delete với DiChuyen.txt và đọc kỹ hộp thoại cảnh báo trước khi xác nhận.", [(23, "Hộp thoại xác nhận xóa vĩnh viễn")]),
                ("Khôi phục tệp", "Mở Recycle Bin, chọn Restore và kiểm tra tệp quay lại đúng vị trí ban đầu.", [(24, "Mở Recycle Bin"), (25, "Chọn Restore"), (26, "Tệp được khôi phục thành công")]),
            ],
        ),
    ]
    sections = []
    toc = []
    for stage_number, title, intro, steps in stages:
        anchor = f"chang-{int(stage_number)}"
        toc.append(f'<a href="#{anchor}"><b>{stage_number}</b><span>{escape(title)}</span></a>')
        step_html = []
        for step_index, (step_title, description, images) in enumerate(steps, 1):
            image_html = "".join(
                f'<button class="tutorial-image" data-src="../assets/b1/image{number}.png" data-caption="Hình {number}. {escape(caption)}"><img src="../assets/b1/image{number}.png" alt="Hình {number}. {escape(caption)}" loading="lazy"><span>Hình {number}</span><small>{escape(caption)}</small></button>'
                for number, caption in images
            )
            step_html.append(
                f'<article class="tutorial-step reveal"><div class="tutorial-copy"><span>{stage_number}.{step_index:02d}</span><h3>{escape(step_title)}</h3><p>{escape(description)}</p></div><div class="tutorial-images">{image_html}</div></article>'
            )
        sections.append(
            f'<section class="tutorial-stage" id="{anchor}"><div class="tutorial-stage-head reveal"><span>Chặng {stage_number}</span><h2>{escape(title)}</h2><p>{escape(intro)}</p></div>{"".join(step_html)}</section>'
        )
    return f"""
    <section class="content-section b1-tutorial-section" id="bao-cao">
      <div class="shell">
        <div class="section-intro reveal"><p class="eyebrow">Hướng dẫn trực quan</p><h2>Thao tác đến đâu,<br><em>xem kết quả đến đó.</em></h2><p>Bài thực hành được trình bày lại thành bốn chặng. Toàn bộ 26 ảnh được đặt đúng cạnh thao tác tương ứng để có thể làm theo mà không phải dò giữa phần chữ và phụ lục.</p></div>
        <nav class="tutorial-map reveal">{"".join(toc)}</nav>
        <div class="compare-note reveal"><div><b>Copy</b><span>Tạo thêm một bản; bản gốc vẫn ở vị trí cũ.</span></div><div><b>Cut</b><span>Chuyển bản gốc sang vị trí mới; vị trí cũ không còn tệp.</span></div><div><b>Delete</b><span>Đưa tệp vào Recycle Bin và có thể Restore.</span></div><div><b>Shift + Delete</b><span>Xóa không qua Recycle Bin, khó phục hồi.</span></div></div>
        <div class="tutorial-flow">{"".join(sections)}</div>
      </div>
    </section>"""


def inline_evidence(project_code, heading):
    groups = INLINE_EVIDENCE.get(project_code, {})
    images = next((items for marker, items in groups.items() if heading.startswith(marker)), None)
    if not images:
        return ""
    figures = "".join(
        f'<button class="context-image" data-src="../{src}" data-caption="{escape(caption)}"><img src="../{src}" alt="{escape(caption)}" loading="lazy"><span>Minh chứng {index:02d}</span><small>{escape(caption)}</small></button>'
        for index, (src, caption) in enumerate(images, 1)
    )
    return f'<div class="inline-evidence reveal"><div class="inline-evidence-head"><span>Ảnh minh chứng tại nội dung này</span><p>Chọn ảnh để xem kích thước lớn.</p></div><div class="inline-evidence-grid count-{len(images)}">{figures}</div></div>'


def online_report(project):
    if project["code"] == "B1":
        return b1_tutorial()
    path = ROOT / "reports" / project["pdf"].replace(".pdf", ".docx")
    document = Document(path)
    style_levels = [
        int(match.group(1))
        for paragraph in document.paragraphs
        if (match := re.search(r"Heading\s+(\d+)", paragraph.style.name))
    ]
    base_style_level = min(style_levels) if style_levels else None
    used = {}
    toc = []
    content = []
    started = False
    numeric_top_level = False
    for block in iter_docx_blocks(document):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if not text:
                continue
            if text.lower().startswith("ghi chú minh chứng:"):
                continue
            level = report_heading(text, block.style.name, base_style_level)
            if not started:
                if level is None:
                    continue
                started = True
                numeric_top_level = bool(re.match(r"^\d+\.\s+", text))
                level = 2
            elif numeric_top_level and re.match(r"^\d+\.\s+", text):
                level = 2
            if text.lower().startswith(("hình ", "bảng ")) and level is None:
                continue
            if level:
                if any(text.startswith(marker) for marker in SKIP_REPORT_HEADINGS.get(project["code"], ())):
                    continue
                anchor = anchor_id(text, used)
                toc.append((level, anchor, text))
                content.append(f'<h{level} id="{anchor}" class="reveal">{escape(text)}</h{level}>')
                evidence = inline_evidence(project["code"], text)
                if evidence:
                    content.append(evidence)
            elif block.style.name.startswith("List"):
                content.append(f'<p class="report-bullet reveal">{escape(text)}</p>')
            else:
                content.append(f'<p class="reveal">{escape(text)}</p>')
        elif started:
            rendered = report_table_html(block)
            if rendered:
                content.append(rendered)
    toc_html = "".join(
        f'<a class="level-{level}" href="#{anchor}">{escape(text)}</a>'
        for level, anchor, text in toc
    )
    return f"""
    <section class="content-section online-report-section" id="bao-cao">
      <div class="shell">
        <div class="section-intro reveal"><p class="eyebrow">Báo cáo trực tuyến</p><h2>Toàn bộ nội dung,<br><em>đọc ngay trên web.</em></h2><p>Nội dung được chuyển từ báo cáo DOCX và tối ưu lại cho màn hình. Bản PDF vẫn được giữ để tải xuống hoặc in.</p></div>
        <div class="report-layout">
          <aside class="report-toc reveal"><span>Mục lục báo cáo</span>{toc_html}<a class="report-pdf" href="../deliverables/Google-Drive-PDF/{project['pdf']}" target="_blank" rel="noopener">Mở bản PDF ↗</a></aside>
          <article class="report-body">{"".join(content)}</article>
        </div>
      </div>
    </section>"""


def heroic_html():
    values = [
        ("H", "Human-in-the-loop", "Con người chịu trách nhiệm cuối cùng."),
        ("E", "Ethical boundaries", "Nhận diện ranh giới đạo đức."),
        ("R", "Reference & fact-check", "Kiểm chứng và tìm nguồn gốc."),
        ("O", "Ownership respect", "Tôn trọng bản quyền và dữ liệu."),
        ("I", "Integrity & transparency", "Khai báo AI trung thực."),
        ("C", "Continuous learning", "Dùng AI để học, không để ngừng học."),
    ]
    return '<div class="heroic-grid">' + "".join(f'<article class="tilt reveal"><b>{a}</b><h3>{escape(b)}</h3><p>{escape(c)}</p></article>' for a, b, c in values) + "</div>"


def project_page(project, index):
    prefix = "../"
    prev_project = PROJECTS[index - 1] if index > 0 else None
    next_project = PROJECTS[index + 1] if index < len(PROJECTS) - 1 else None
    prev_link = f'<a href="{prev_project["slug"]}.html"><span>← Bài trước</span><strong>{escape(prev_project["title"])}</strong></a>' if prev_project else '<a href="../index.html"><span>← Quay lại</span><strong>Trang chủ portfolio</strong></a>'
    next_link = f'<a class="next" href="{next_project["slug"]}.html"><span>Bài tiếp theo →</span><strong>{escape(next_project["title"])}</strong></a>' if next_project else '<a class="next" href="tong-ket.html"><span>Đi đến →</span><strong>Tổng kết học phần</strong></a>'
    special = heroic_html() if project.get("special") == "heroic" else ""
    table = table_html(project.get("table"))
    body = f"""
    <section class="case-hero" data-code="{project['code']}">
      <div class="case-orb one"></div><div class="case-orb two"></div>
      <div class="shell case-hero-grid">
        <div class="case-title reveal">
          <a class="back-link" href="../index.html#projects">← Tất cả bài tập</a>
          <p class="eyebrow">{escape(project['category'])} · {project['code']}</p>
          <h1>{escape(project['title'])}</h1>
          <p class="case-tagline">{escape(project['tagline'])}</p>
          <div class="hero-actions"><a class="button primary magnetic" href="#bao-cao">Đọc báo cáo trên web</a><a class="button ghost magnetic" href="../deliverables/Google-Drive-PDF/{project['pdf']}" target="_blank" rel="noopener">Mở báo cáo PDF ↗</a></div>
        </div>
        <div class="case-index reveal"><span>{project['code']}</span><small>{index + 1:02d} / 07</small></div>
      </div>
      <div class="shell metric-strip">{metrics_html(project['metrics'])}</div>
    </section>
    <section class="content-section" id="story">
      <div class="shell story-grid">
        <aside class="story-aside reveal"><p class="eyebrow">Câu chuyện bài tập</p><h2>Từ yêu cầu đến<br><em>năng lực thực tế.</em></h2><div class="aside-line"></div><p>{escape(project['tagline'])}</p></aside>
        <div class="story-copy reveal"><p class="lead">{escape(project['summary'])}</p><div class="objective-list">{"".join(f'<div><span>0{i}</span><p>{escape(item)}</p></div>' for i, item in enumerate(project['objectives'], 1))}</div></div>
      </div>
    </section>
    <section class="content-section tint-section">
      <div class="shell"><div class="section-intro reveal"><p class="eyebrow">Quy trình</p><h2>Từng bước giải quyết<br><em>một vấn đề cụ thể.</em></h2></div><div class="insight-grid numbered">{cards(project['process'], True)}</div></div>
    </section>
    <section class="content-section">
      <div class="shell"><div class="section-intro reveal"><p class="eyebrow">Kết quả & phân tích</p><h2>Điều nổi bật sau<br><em>quá trình thực hiện.</em></h2></div>{table}<div class="insight-grid">{cards(project['results'])}</div>{special}</div>
    </section>
    {online_report(project)}
    <section class="content-section lesson-section" id="bai-hoc"><div class="shell lesson-card reveal"><span>Bài học cá nhân</span><blockquote>{escape(project['lesson'])}</blockquote><a class="button primary magnetic" href="../deliverables/Google-Drive-PDF/{project['pdf']}" target="_blank" rel="noopener">Đọc báo cáo đầy đủ ↗</a></div></section>
    <nav class="case-pagination shell">{prev_link}{next_link}</nav>
    """
    return doc(f"{project['code']} · {project['title']}", body, prefix, "projects", project["summary"])


def reflection_page():
    body = """
    <section class="reflection-hero"><div class="aurora a1"></div><div class="aurora a2"></div><div class="shell"><p class="eyebrow reveal">Tổng kết học phần · 2026</p><h1 class="reveal">Từ “biết dùng”<br>đến <em>biết làm chủ.</em></h1><p class="reveal">Cảm nghĩ sau hành trình trải nghiệm công nghệ số và trí tuệ nhân tạo.</p></div></section>
    <section class="content-section"><div class="shell reflection-story"><aside class="story-aside reveal"><p class="eyebrow">Điểm xuất phát</p><h2>Ban đầu, tôi nghĩ đây là môn học về <em>công cụ.</em></h2></aside><div class="story-copy reveal"><p class="lead">Trước khi học, tôi thường đánh giá năng lực số bằng việc một người sử dụng được bao nhiêu phần mềm. Sau bảy bài tập, tôi nhận ra công cụ chỉ là phần dễ nhìn thấy nhất.</p><p>Năng lực thực sự nằm ở cách tổ chức dữ liệu trước khi bắt đầu, cách tìm và kiểm chứng thông tin trước khi tin, cách giao tiếp yêu cầu để AI hiểu đúng, cách phối hợp với người khác và cách chịu trách nhiệm về sản phẩm cuối cùng.</p><p>Điều này khiến tôi nhìn những thao tác rất cơ bản, như đặt tên tệp hay ghi chú nguồn tài liệu, bằng một thái độ khác. Chúng không phải việc phụ; chúng là nền móng để những công việc phức tạp hơn không bị hỗn loạn.</p></div></div></section>
    <section class="content-section tint-section"><div class="shell"><div class="section-intro reveal"><p class="eyebrow">Bảy thay đổi</p><h2>Mỗi bài tập để lại<br><em>một thay đổi nhỏ.</em></h2></div><div class="change-grid">
      <article class="reveal"><b>B1</b><h3>Tổ chức trước khi làm</h3><p>Tôi chú ý hơn đến cấu trúc thư mục, tên tệp và khả năng phục hồi dữ liệu.</p></article>
      <article class="reveal"><b>B2</b><h3>Kiểm chứng trước khi tin</h3><p>Tôi không còn xem kết quả tìm kiếm đầu tiên là câu trả lời cuối cùng.</p></article>
      <article class="reveal"><b>B3</b><h3>Hỏi rõ trước khi đòi câu trả lời tốt</h3><p>Tôi học cách biến yêu cầu mơ hồ thành tiêu chí cụ thể và kiểm tra được.</p></article>
      <article class="reveal"><b>B4</b><h3>Minh bạch khi cộng tác</h3><p>Tiến độ, trách nhiệm và quyết định cần được ghi lại ở đúng nơi.</p></article>
      <article class="reveal"><b>B5</b><h3>Chọn lọc trước khi sáng tạo</h3><p>AI tạo ra nhiều phương án; bản sắc đến từ quyết định của con người.</p></article>
      <article class="reveal"><b>B6</b><h3>Chịu trách nhiệm khi dùng AI</h3><p>Minh bạch và kiểm chứng không phải giới hạn, mà là tiêu chuẩn chất lượng.</p></article>
      <article class="reveal"><b>B7</b><h3>Dùng AI để mở rộng, không thay thế tư duy</h3><p>Công cụ nghiên cứu giúp đi nhanh, nhưng hướng đi vẫn do người dùng quyết định.</p></article>
    </div></div></section>
    <section class="content-section dark-section"><div class="shell reflection-quote reveal"><span>Điều tôi nhớ nhất</span><blockquote>“AI không khiến một sản phẩm tự động trở nên tốt. Nó chỉ làm tốc độ tạo ra cả phương án tốt lẫn phương án tệ nhanh hơn. Năng lực của người dùng nằm ở việc phân biệt hai điều đó.”</blockquote></div></section>
    <section class="content-section"><div class="shell future-grid"><div class="section-intro reveal"><p class="eyebrow">Nhìn về phía trước</p><h2>Học phần kết thúc.<br><em>Thói quen thì ở lại.</em></h2></div><div class="future-list reveal"><div><span>01</span><p>Tiếp tục duy trì hệ thống quản lý tài liệu và nguồn tham khảo cho các môn học sau.</p></div><div><span>02</span><p>Sử dụng prompt như một bản đặc tả nhiệm vụ, không phải câu lệnh xin đáp án.</p></div><div><span>03</span><p>Đưa bước kiểm chứng và khai báo AI thành mặc định trong mọi sản phẩm học thuật.</p></div><div><span>04</span><p>Thử nghiệm công cụ mới với tinh thần tò mò, nhưng chỉ giữ lại thứ tạo ra giá trị thực.</p></div></div></div></section>
    <section class="final-note content-section"><div class="shell reveal"><p class="eyebrow">Lời kết</p><h2>Tôi không kết thúc học phần với cảm giác đã “biết AI”.<br><em>Tôi kết thúc với một cách học tốt hơn.</em></h2><p>Công nghệ sẽ tiếp tục thay đổi. Điều có thể duy trì lâu dài là khả năng tự học, tư duy phản biện, giao tiếp rõ ràng và trách nhiệm với lựa chọn của mình.</p><a class="button primary magnetic" href="../index.html#projects">Xem lại bảy bài tập →</a></div></section>
    """
    return doc("Tổng kết học phần", body, "../", "reflection", "Cảm nghĩ sau học phần Công nghệ số và Ứng dụng Trí tuệ nhân tạo.")


def main():
    for index, project in enumerate(PROJECTS):
        (PAGES / f"{project['slug']}.html").write_text(project_page(project, index), encoding="utf-8")
    (PAGES / "tong-ket.html").write_text(reflection_page(), encoding="utf-8")
    print(f"Built {len(PROJECTS) + 1} detail pages; kept the original index.html")


if __name__ == "__main__":
    main()
