import React, { useState } from "react";
import axios from "axios";

const FileUpload = () => {
    const [file, setFile] = useState(null);
    const [message, setMessage] = useState("");

    const handleFileChange = (event) => {
        setFile(event.target.files[0]);
    };

    const handleUpload = async () => {
        if (!file) {
            setMessage("Please select a file!");
            return;
        }

        const formData = new FormData();
        formData.append("file", file);

        try {
            const response = await axios.post("http://localhost:3000/upload", formData, {
                headers: { "Content-Type": "multipart/form-data" },
            });

            setMessage(response.data.message);
        } catch (error) {
            setMessage("Error uploading file.");
        }
    };

    return (
        <div className="file-upload">
            <h2>Upload Excel File</h2>
            <input type="file" accept=".xlsx" onChange={handleFileChange} />
            <button className="btn upload-btn" onClick={handleUpload}>Upload</button>
            {message && <p>{message}</p>}
        </div>
    );
};

export default FileUpload;
