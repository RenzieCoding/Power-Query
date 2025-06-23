![Frame 1](https://github.com/user-attachments/assets/4ca99b3c-da89-42a8-8c7e-3158e9aefefb)
# Combining Multiple Excel Files Using Power Query (M Language)

This Power Query solution handles the challenge of combining multiple Excel files from a folder — each with **inconsistent formats**, **different sheet/tab names**, and **month headers in other languages**.

Some tables start on different rows or have non-standard headers, making this a tricky transformation task. This solution dynamically cleans and consolidates all files into one structured table.

---

## 🛠️ What This Does

✔️ Connects to a folder of Excel files  
✔️ Dynamically identifies and processes multiple sheets  
✔️ Removes unnecessary rows and promotes headers  
✔️ Converts month names in different languages  
✔️ Combines everything into one clean, ready-to-use table

---

## 📦 Files Included

- `CombineExcelFiles.pq` – the full M code script  
- `workflow.png` – visual breakdown of the Power Query steps *(optional)*  
- `sample-files/` – sample folder structure or dummy Excel files *(if applicable)*

---

## 📚 Techniques Used

- Record, Table, and List manipulation  
- `Table.Skip`, `Table.PromoteHeaders`, and dynamic filtering  
- Folder and Sheet iteration  
- Applied Steps broken into reusable blocks

---

## 🙌 Credits

- Based on ideas and techniques from  
  📘 *Power Query: Beyond the User Interface* by **Chandeep Chhabra**  
  💡 Tips and guidance from **Pedro Bagtas**, senior M ninja 🥷

---

## 📎 LinkedIn Post

You can view the original post and visual walkthrough here:  
🔗 []

---

## 📬 Questions?

Feel free to connect or open an issue if you have questions or want to collaborate!
