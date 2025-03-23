import React from "react";
import FileUpload from "./components/FileUpload";
import ProcessFile from "./components/ProcessFile";
import "./App.css";
import logo1 from "./assets/logo1.jpg"; // Ensure the correct path
import logo2 from "./assets/logo2.jpg"; // Ensure the correct path

const App = () => {
    return (
        <div>
            {/* Navbar */}
            <nav className="navbar">
                <img src={logo1} alt="College Logo 1" className="nav-logo" />
                <h1 className="college-name">
                    JSPM'S Jaywantrao Sawant College of Engineering, Hadapsar, Pune - 411028
                </h1>
                <img src={logo2} alt="College Logo 2" className="nav-logo" />
            </nav>

            {/* Main Content */}
            <div className="container">
                <h1 className="main-heading">OBE Sheet Creator</h1>
                
                {/* Download Template Button */}
                <a href="/template.xlsx" download>
                    <button className="btn template-btn">Download Template</button>
                </a>

                <FileUpload />
                <ProcessFile />
            </div>
        </div>
    );
};

export default App;
