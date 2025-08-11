
import Foundation

// OLE2/CFB reader (to be implemented in Step 2)
struct CFBFile {
    let data: Data
    init(fileURL: URL) throws {
        self.data = try Data(contentsOf: fileURL)
        // Step 2: validate header, build FAT/MiniFAT, directory, etc.
    }

    func stream(named: String) throws -> Data {
        // Step 2: return the bytes of "Workbook" or "Book"
        throw XLSReadError.workbookStreamMissing
    }
}

//git add .
//git commit -m "chore: initial SPM scaffold for ClassicXLS"
//git branch -M main
//# If you use GitHub CLI:
//gh repo create youruser/ClassicXLS --public --source=. --push
//# Or create an empty GitHub repo manually and:
//git remote add origin git@github.com:youruser/ClassicXLS.git
//git push -u origin main
