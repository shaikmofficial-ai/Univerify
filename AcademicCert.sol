// SPDX-License-Identifier: MIT
pragma solidity ^0.8.0;

contract AcademicCertificates {
    // This creates a "record" for a certificate
    struct Certificate {
        string studentName;
        string courseName;
        string ipfsHash; // This links to the actual PDF file
    }

    // A list of all certificates issued
    mapping(uint256 => Certificate) public certificates;
    uint256 public nextId;

    // The "University" (you) who owns this contract
    address public university;

    constructor() {
        university = msg.sender;
    }

    // Function to issue a new degree
    function issueDegree(string memory _name, string memory _course, string memory _hash) public {
        require(msg.sender == university, "Only the University can issue degrees!");
        certificates[nextId] = Certificate(_name, _course, _hash);
        nextId++;
    }
}