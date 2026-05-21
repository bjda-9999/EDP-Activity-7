-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1:3306
-- Generation Time: May 13, 2026 at 11:29 AM
-- Server version: 10.4.32-MariaDB
-- PHP Version: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `school`
--

DELIMITER $$
--
-- Procedures
--
CREATE DEFINER=`root`@`localhost` PROCEDURE `spgetstudentgrades` (IN `stuid` INT)   BEGIN
    SELECT c.coursename, e.grade
    FROM Enrollment e
    JOIN Course c ON e.courseid = c.courseid
    WHERE e.studentid = stuid;
END$$

--
-- Functions
--
CREATE DEFINER=`root`@`localhost` FUNCTION `fncalculategpa` (`stuid` INT) RETURNS DECIMAL(3,2) DETERMINISTIC BEGIN
    DECLARE gpa DECIMAL(3,2);

    SELECT AVG(
        CASE grade
            WHEN 'A' THEN 4.00
            WHEN 'B' THEN 3.00
            WHEN 'C' THEN 2.00
            WHEN 'D' THEN 1.00
            WHEN 'F' THEN 0.00
        END
    ) INTO gpa
    FROM Enrollment
    WHERE studentid = stuid;

    RETURN gpa;
END$$

DELIMITER ;

-- --------------------------------------------------------

--
-- Table structure for table `accounts`
--

CREATE TABLE `accounts` (
  `account_id` int(11) NOT NULL,
  `username` varchar(100) NOT NULL,
  `full_name` varchar(150) NOT NULL,
  `email` varchar(150) NOT NULL,
  `password_hash` varchar(255) NOT NULL,
  `role` enum('Administrator','Staff','Viewer') NOT NULL DEFAULT 'Staff',
  `security_question` varchar(255) DEFAULT NULL,
  `security_answer` varchar(255) DEFAULT NULL,
  `is_active` tinyint(1) NOT NULL DEFAULT 1,
  `created_at` datetime NOT NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `accounts`
--

INSERT INTO `accounts` (`account_id`, `username`, `full_name`, `email`, `password_hash`, `role`, `security_question`, `security_answer`, `is_active`, `created_at`) VALUES
(1, 'admin', 'System Administrator', 'admin@school.local', 'admin123', 'Administrator', 'What is your mother\'s maiden name?', 'dela Cruz', 1, '2026-05-09 20:08:01'),
(2, 'benj', 'Benjamin D. Aguilar Jr.', 'benj@gmail.com', '123Jax', 'Staff', 'What was the name of your first pet?', 'Jax', 1, '2026-05-09 20:09:16');

-- --------------------------------------------------------

--
-- Table structure for table `course`
--

CREATE TABLE `course` (
  `courseid` int(11) NOT NULL,
  `coursename` varchar(100) NOT NULL,
  `deptid` int(11) DEFAULT NULL,
  `instructorid` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `course`
--

INSERT INTO `course` (`courseid`, `coursename`, `deptid`, `instructorid`) VALUES
(1, 'Database Systems', 1, 1),
(2, 'Data Structures', 1, 2),
(3, 'Web Development', 2, 3),
(4, 'Marketing 101', 3, 4),
(5, 'Thermodynamics', 5, 7),
(6, 'Nursing Care', 6, 9),
(7, 'Teaching Methods', 7, 2),
(8, 'Abnormal Psychology', 8, 8),
(9, 'Criminal Law', 9, 4),
(10, 'Architectural Design', 10, 10);

-- --------------------------------------------------------

--
-- Table structure for table `department`
--

CREATE TABLE `department` (
  `deptid` int(11) NOT NULL,
  `deptname` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `department`
--

INSERT INTO `department` (`deptid`, `deptname`) VALUES
(1, 'Computer Science'),
(2, 'Information Technology'),
(3, 'Business Administration'),
(4, 'Accounting'),
(5, 'Engineering'),
(6, 'Nursing'),
(7, 'Education'),
(8, 'Psychology'),
(9, 'Criminology'),
(10, 'Architecture');

-- --------------------------------------------------------

--
-- Table structure for table `enrollment`
--

CREATE TABLE `enrollment` (
  `enrollmentid` int(11) NOT NULL,
  `studentid` int(11) DEFAULT NULL,
  `courseid` int(11) DEFAULT NULL,
  `grade` char(1) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `enrollment`
--

INSERT INTO `enrollment` (`enrollmentid`, `studentid`, `courseid`, `grade`) VALUES
(1, 1, 1, 'A'),
(2, 2, 2, 'B'),
(3, 3, 3, 'A'),
(4, 4, 4, 'C'),
(5, 5, 5, 'B'),
(6, 6, 6, 'A'),
(7, 7, 7, 'B'),
(8, 8, 8, 'C'),
(9, 9, 9, 'B'),
(10, 10, 10, 'A');

--
-- Triggers `enrollment`
--
DELIMITER $$
CREATE TRIGGER `trg_before_delete_enrollment` BEFORE DELETE ON `enrollment` FOR EACH ROW BEGIN
    IF OLD.grade = 'A' THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'Cannot delete enrollment with grade A.';
    END IF;
END
$$
DELIMITER ;
DELIMITER $$
CREATE TRIGGER `trg_before_insert_enrollment` BEFORE INSERT ON `enrollment` FOR EACH ROW BEGIN
    IF NEW.grade NOT IN ('A','B','C','D','F') THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'Invalid Grade! Must be A, B, C, D, or F.';
    END IF;
END
$$
DELIMITER ;
DELIMITER $$
CREATE TRIGGER `trg_before_update_enrollment` BEFORE UPDATE ON `enrollment` FOR EACH ROW BEGIN
    IF NEW.grade NOT IN ('A','B','C','D','F') THEN
        SIGNAL SQLSTATE '45000'
        SET MESSAGE_TEXT = 'Invalid Grade Update!';
    END IF;
END
$$
DELIMITER ;

-- --------------------------------------------------------

--
-- Table structure for table `instructor`
--

CREATE TABLE `instructor` (
  `instructorid` int(11) NOT NULL,
  `firstname` varchar(50) NOT NULL,
  `lastname` varchar(50) NOT NULL,
  `email` varchar(100) DEFAULT NULL,
  `deptid` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `instructor`
--

INSERT INTO `instructor` (`instructorid`, `firstname`, `lastname`, `email`, `deptid`) VALUES
(1, 'Alan', 'Turing', 'alan@email.com', 1),
(2, 'Grace', 'Hopper', 'grace@email.com', 1),
(3, 'Steve', 'Jobs', 'steve@email.com', 2),
(4, 'Bill', 'Gates', 'bill@email.com', 3),
(5, 'Elon', 'Musk', 'elon@email.com', 5),
(6, 'Marie', 'Curie', 'marie@email.com', 6),
(7, 'Albert', 'Einstein', 'albert@email.com', 5),
(8, 'Sigmund', 'Freud', 'sigmund@email.com', 8),
(9, 'Florence', 'Nightingale', 'florence@email.com', 6),
(10, 'Frank', 'Lloyd', 'frank@email.com', 10);

-- --------------------------------------------------------

--
-- Table structure for table `student`
--

CREATE TABLE `student` (
  `studentid` int(11) NOT NULL,
  `firstname` varchar(50) NOT NULL,
  `lastname` varchar(50) NOT NULL,
  `email` varchar(100) DEFAULT NULL,
  `deptid` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `student`
--

INSERT INTO `student` (`studentid`, `firstname`, `lastname`, `email`, `deptid`) VALUES
(1, 'John', 'Doe', 'john@email.com', 1),
(2, 'Jane', 'Smith', 'jane@email.com', 1),
(3, 'Mark', 'Lee', 'mark@email.com', 2),
(4, 'Anna', 'Taylor', 'anna@email.com', 3),
(5, 'Paul', 'Brown', 'paul@email.com', 4),
(6, 'Lisa', 'White', 'lisa@email.com', 5),
(7, 'Tom', 'Harris', 'tom@email.com', 6),
(8, 'Emma', 'Clark', 'emma@email.com', 7),
(9, 'Chris', 'Lewis', 'chris@email.com', 8),
(10, 'Sophia', 'Walker', 'sophia@email.com', 9);

-- --------------------------------------------------------

--
-- Stand-in structure for view `vwdepartmentsummary`
-- (See below for the actual view)
--
CREATE TABLE `vwdepartmentsummary` (
`deptname` varchar(100)
,`totalinstructors` bigint(21)
,`totalcourses` bigint(21)
);

-- --------------------------------------------------------

--
-- Stand-in structure for view `vwinstructorload`
-- (See below for the actual view)
--
CREATE TABLE `vwinstructorload` (
`instructorid` int(11)
,`instructorname` varchar(101)
,`totalcourses` bigint(21)
);

-- --------------------------------------------------------

--
-- Stand-in structure for view `vwstudentenrollments`
-- (See below for the actual view)
--
CREATE TABLE `vwstudentenrollments` (
`studentid` int(11)
,`studentname` varchar(101)
,`coursename` varchar(100)
,`grade` char(1)
);

-- --------------------------------------------------------

--
-- Structure for view `vwdepartmentsummary`
--
DROP TABLE IF EXISTS `vwdepartmentsummary`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vwdepartmentsummary`  AS SELECT `d`.`deptname` AS `deptname`, count(distinct `i`.`instructorid`) AS `totalinstructors`, count(distinct `c`.`courseid`) AS `totalcourses` FROM ((`department` `d` left join `instructor` `i` on(`d`.`deptid` = `i`.`deptid`)) left join `course` `c` on(`d`.`deptid` = `c`.`deptid`)) GROUP BY `d`.`deptname` ;

-- --------------------------------------------------------

--
-- Structure for view `vwinstructorload`
--
DROP TABLE IF EXISTS `vwinstructorload`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vwinstructorload`  AS SELECT `i`.`instructorid` AS `instructorid`, concat(`i`.`firstname`,' ',`i`.`lastname`) AS `instructorname`, count(`c`.`courseid`) AS `totalcourses` FROM (`instructor` `i` left join `course` `c` on(`i`.`instructorid` = `c`.`instructorid`)) GROUP BY `i`.`instructorid` ;

-- --------------------------------------------------------

--
-- Structure for view `vwstudentenrollments`
--
DROP TABLE IF EXISTS `vwstudentenrollments`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vwstudentenrollments` AS 
SELECT `s`.`studentid` AS `studentid`, 
       CONCAT(`s`.`lastname`, ', ', `s`.`firstname`) AS `studentname`, 
       GROUP_CONCAT(DISTINCT `c`.`coursename` SEPARATOR '; ') AS `courses`,
       GROUP_CONCAT(DISTINCT CONCAT(`c`.`coursename`, ': ', `e`.`grade`) SEPARATOR ' | ') AS `grades_summary`
FROM `student` `s`
JOIN `enrollment` `e` ON `s`.`studentid` = `e`.`studentid`
JOIN `course` `c` ON `e`.`courseid` = `c`.`courseid`
GROUP BY `s`.`studentid`;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `accounts`
--
ALTER TABLE `accounts`
  ADD PRIMARY KEY (`account_id`),
  ADD UNIQUE KEY `uq_username` (`username`);

--
-- Indexes for table `course`
--
ALTER TABLE `course`
  ADD PRIMARY KEY (`courseid`),
  ADD KEY `deptid` (`deptid`),
  ADD KEY `instructorid` (`instructorid`);

--
-- Indexes for table `department`
--
ALTER TABLE `department`
  ADD PRIMARY KEY (`deptid`);

--
-- Indexes for table `enrollment`
--
ALTER TABLE `enrollment`
  ADD PRIMARY KEY (`enrollmentid`),
  ADD KEY `studentid` (`studentid`),
  ADD KEY `courseid` (`courseid`),
  ADD UNIQUE KEY `uq_student_course` (`studentid`, `courseid`);

--
-- Indexes for table `instructor`
--
ALTER TABLE `instructor`
  ADD PRIMARY KEY (`instructorid`),
  ADD KEY `deptid` (`deptid`);

--
-- Indexes for table `student`
--
ALTER TABLE `student`
  ADD PRIMARY KEY (`studentid`),
  ADD KEY `deptid` (`deptid`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `accounts`
--
ALTER TABLE `accounts`
  MODIFY `account_id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- AUTO_INCREMENT for table `course`
--
ALTER TABLE `course`
  MODIFY `courseid` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;

--
-- AUTO_INCREMENT for table `department`
--
ALTER TABLE `department`
  MODIFY `deptid` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;

--
-- AUTO_INCREMENT for table `enrollment`
--
ALTER TABLE `enrollment`
  MODIFY `enrollmentid` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=14;

--
-- AUTO_INCREMENT for table `instructor`
--
ALTER TABLE `instructor`
  MODIFY `instructorid` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;

--
-- AUTO_INCREMENT for table `student`
--
ALTER TABLE `student`
  MODIFY `studentid` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;

--
-- Constraints for dumped tables
--

--
-- Constraints for table `course`
--
ALTER TABLE `course`
  ADD CONSTRAINT `course_ibfk_1` FOREIGN KEY (`deptid`) REFERENCES `department` (`deptid`),
  ADD CONSTRAINT `course_ibfk_2` FOREIGN KEY (`instructorid`) REFERENCES `instructor` (`instructorid`);

--
-- Constraints for table `enrollment`
--
ALTER TABLE `enrollment`
  ADD CONSTRAINT `enrollment_ibfk_1` FOREIGN KEY (`studentid`) REFERENCES `student` (`studentid`),
  ADD CONSTRAINT `enrollment_ibfk_2` FOREIGN KEY (`courseid`) REFERENCES `course` (`courseid`);

--
-- Constraints for table `instructor`
--
ALTER TABLE `instructor`
  ADD CONSTRAINT `instructor_ibfk_1` FOREIGN KEY (`deptid`) REFERENCES `department` (`deptid`);

--
-- Constraints for table `student`
--
ALTER TABLE `student`
  ADD CONSTRAINT `student_ibfk_1` FOREIGN KEY (`deptid`) REFERENCES `department` (`deptid`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
