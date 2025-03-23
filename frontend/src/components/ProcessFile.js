import React, { useState } from "react";
import axios from "axios";

const ProcessFile = () => {
    const [downloadLink, setDownloadLink] = useState("");
    const [message, setMessage] = useState("");

    const handleProcess = async () => {
        try {
            const response = await axios.get("http://localhost:3000/process");
            setDownloadLink("http://localhost:3000/processed_files/Processed_File.xlsx");
            setMessage(response.data.message);
        } catch (error) {
            setMessage("Error processing file.");
        }
    };

    return (
        <div className="file-process">
            <h2>Process Excel File</h2>
            <button className="btn process-btn" onClick={handleProcess}>Process File</button>
            {message && <p>{message}</p>}
            {downloadLink && <a href={downloadLink} download>Download Processed File</a>}
        </div>
    );
};

export default ProcessFile;
