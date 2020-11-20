// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2020  Mozard BV
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

import React from "react";
import styled from "styled-components";
import { v4 as uuidv4 } from "uuid";

const StyledLogo = styled.div`
  width: 100px;
  height: 100px;
  overflow: hidden;
  background-color: hsl(0, 0%, 100%);
  border-radius: 9999px;
`;

function Logo() {
  const _descId = uuidv4();
  const _titleId = uuidv4();

  return (
    <StyledLogo>
      <svg
        aria-describedby={_descId}
        aria-labelledby={_titleId}
        xmlns="http://www.w3.org/2000/svg"
        version="1.1"
        viewBox="0 0 61.1 61.2"
      >
        <title id={_titleId}>Mozard</title>
        <desc id={_descId}>Logo van het bedrijf Mozard</desc>
        <path fill="none" d="M-102.2-119.4h470v300h-470z" />
        <path
          d="M35.2 28.2c.3 0 .6.1.8.1 0 0-.6-.9-1-1-.3 0-.6.1-.9.2-.5.2-1 .4-1.4.8 0 .1 0 .3.1.4.8-.1 1.6-.3 2.4-.5z"
          className="logo-cls-1"
        />
        <path
          d="M5.7 43.3c.9-.5 1.6-1.7 2.8-2.3.9-.5 2.8-.3 3.2-.8s.7-1 1-1.5c1.1-.7 2.6-1.6 3.4-2.1 1.3-.7 2.6-1.5 3.8-2.4.3-.2.3-.5.9-.7v.1c-.8.8-.1 2 .4 3 .9 1.3 2 2.4 3.2 3.3.6.5 1.1 1.1 1.6 1s1-.9 1.2-1.1c2.8-1.9 5.6-3.7 8.5-5.3 1.2-.6 2.3-.5 3.4-.9.4-.3 14.8-22.6 21.3-33.6H0v44.2c1.2-.2 5.3-.8 5.7-.9zm16-10.2c.1-.2.5-.4.7-.6.4-.5.7-1 1-1.6.4-1.2.9-2.3 1.3-3.4.4-.8.9-1.6 1.4-2.4.9-1.1 1.9-2.2 2.8-3.3l.6-.9c.6-.4 1.2-.8 1.9-1.1l1.5-1.2c.4-.2 1.5-.9 2-.7.2 0 .4.1.5.3.5 1-.5 1.5-1.1 2-.9.6-1.7 1.3-2.5 2.1-.2.3-1.1 1.3-1.2 1.5.2.2 1.6-1.1 2.4-1.6.4-.2.9-.2 1.3-.1.3 0 3.6.2 3.8.2s.5.7.5.9c0 .3-.1.5-.3.6-.2.1-3.1.3-3.4.4-.2 0-.3.1-.4.2-.1.1-.8.5-1.2.8s-.7.4-.8.8c0 .1.5.2.7.1.6-.3 1.3-.5 1.9-.7.3.1.7.3.9.5l.6.6c.5.7.9 1.4 1.2 2.2 0 .1 0 .4-.2.4s-.4.1-.6.2c.2.6.2 1.2-.2 1.7-.3.1-.7.1-1-.2-.2-.2-.3-.4-.3-.6-.2.1-2 1.3-2.2 1.5-.3.3-.5.6-.7.9-.3.4-.6.9-1 1.3-.7.7-1.4 1.3-2.1 1.8-.5.4-2.8 2.6-2.9 2.7-.6-.2-1.2-.5-1.7-.9-.3-.2-.7-.4-1-.6-.4-.3-.8-.7-1.1-1.1-.5-.4-.8-.9-1-1.5-.1-.2-.1-.4-.2-.7-.2-.2-.1-.4.1-.5z"
          className="logo-cls-1"
        />
        <path
          fill="#717075"
          d="M42.2 33.5c.2.1.3.2.5.4.2.5.3.9.3 1.4.6-.3 1.3-.6 2-.7.5-.1 1.1 0 1.5.4.2.3.2.6.2.9-.4.4-.8.7-1.3.9L42.8 38l-1.7.9c-.3.2-.5.4-.4.5s.9 0 1.3-.2c.7-.2 3.5-1.5 4.3-2 .4-.2.9-.4 1.4-.5.4 0 .8.4.9.8-.1.5-.4.9-.7 1.2-1 .7-2 1.4-3.1 1.9-.9.3-3.2.7-3.7 1.4-.1.1-.1.3-.1.4.5.1.9.2 1.4.2.6 0 1.3 0 1.9-.2.8-.3 1.7-.5 2.5-.7.4 0 .7.2 1 .4.1.3.1.3 0 .5s-3.1 1.4-3.7 1.7c-1.3.6-3 .3-4.6 1.1-1.6.8-3.3 1.6-5 2.2-.7.2-1.5.2-2.2.4-.4.1-.7.3-1.1.4-.3-1.4-.8-2.7-1.6-3.8-.5-.7-1.2-1.2-1.9-1.7-.5-.4-1.2-.8-1.8-.9-.4.1-.7.3-1 .5v.2c.7.4 1.4.9 1.8 1.6 1 1.2 1.9 2.5 2.6 3.9.5 1 2 2.6 1.8 3-2.1 3.6-8.9 9.2-9.7 9.6l-.3.3H61V.3L42.2 33.5z"
        />
      </svg>
    </StyledLogo>
  );
}

export default Logo;
